import { NextRequest, NextResponse } from "next/server";
import crypto from "crypto";

/**
 * 슬랙 웹훅 수신 엔드포인트
 *
 * 슬랙 Event Subscriptions에서 이 URL을 등록하면:
 * 1. 지정된 채널에 메시지가 올 때 슬랙이 여기로 POST를 보냄
 * 2. 마감기한 파싱 → 없으면 스레드로 되묻기 → 확정 시 대시보드 등록
 *
 * 필요한 환경변수 (Vercel > Settings > Environment Variables):
 *   SLACK_SIGNING_SECRET  — Slack 앱 > Basic Information > Signing Secret
 *   SLACK_BOT_TOKEN       — Slack 앱 > OAuth & Permissions > Bot User OAuth Token
 *
 * 슬랙 앱 설정:
 *   Event Subscriptions > Request URL: https://<vercel-domain>/api/slack-webhook
 *   Subscribe to bot events: message.channels, message.groups
 *   OAuth Scopes: chat:write, channels:history, groups:history
 *
 * 모니터링 채널:
 *   #section-전략추진실-all (C08NNP1D3A9)
 *   #section-전략추진실-창준님 (C0AJ265GP8W)
 *   #wg-전략추진실 (C09L1LBK1GD)
 */

// ─── 인메모리 스토어 (Vercel 인스턴스 내 유지) ───

interface SlackTask {
  id: string;
  from: string;
  to: string;
  message: string;
  channel: string;
  timestamp: string;
  deadline: string;
  slackTs: string;
}

interface PendingTask extends Omit<SlackTask, "deadline"> {}

/** 마감기한 확정된 태스크 */
const confirmedTasks: SlackTask[] = [];

/** 삭제된 태스크 ID 목록 (프론트 폴링 시 전달 후 비움) */
const deletedTaskIds: string[] = [];

/**
 * 마감기한 미확인 태스크 (스레드 답장 대기 중)
 * key: 원본 메시지의 ts (스레드 루트)
 */
const awaitingDeadline = new Map<string, PendingTask>();

// ─── 유저/채널 매핑 ───

const USER_MAP: Record<string, string> = {
  U08JWR295EC: "이창준",
  U0A26MKAP7E: "주호연",
  U0A7J7J437A: "임성욱",
  U0ALAGP8E79: "임한솔",
  U09A4D5KFUH: "나여준",
  U097N6HRYER: "김범석",
  U09LXJNF99N: "방지수",
  U09L3B29490: "박관우",
};

const CHANNEL_MAP: Record<string, string> = {
  C08NNP1D3A9: "#section-전략추진실-all",
  C0AJ265GP8W: "#section-전략추진실-창준님",
  C09L1LBK1GD: "#wg-전략추진실",
  C0AL18T5KU6: "#wg-사업성장팀x전략추진실",
};

// ─── 유틸 함수 ───

/** 첫 번째 @멘션 유저 추출 */
function extractMentionedUser(text: string): { id: string; name: string } | null {
  const match = text.match(/<@(U[A-Z0-9]+)>/);
  if (!match) return null;
  return { id: match[1], name: USER_MAP[match[1]] || match[1] };
}

/** 슬랙 마크업 제거 */
function cleanSlackText(text: string): string {
  return text
    .replace(/<@U[A-Z0-9]+>/g, "")
    .replace(/<#C[A-Z0-9]+\|([^>]+)>/g, "#$1")
    .replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, "$2")
    .replace(/<(https?:\/\/[^>]+)>/g, "$1")
    .trim();
}

/**
 * 텍스트에서 마감기한 추출
 * 인식 패턴: 오늘/내일/모레, 이번주|다음주 N요일, N월 N일, YYYY.MM.DD, 월말
 * @returns ISO date string (YYYY-MM-DD) 또는 null
 */
function parseDeadline(text: string): string | null {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const toISO = (d: Date) => d.toISOString().split("T")[0];

  if (/오늘/.test(text)) return toISO(now);

  if (/내일/.test(text)) {
    const d = new Date(now);
    d.setDate(d.getDate() + 1);
    return toISO(d);
  }

  if (/모레/.test(text)) {
    const d = new Date(now);
    d.setDate(d.getDate() + 2);
    return toISO(d);
  }

  const dayMap: Record<string, number> = {
    월: 1, 화: 2, 수: 3, 목: 4, 금: 5, 토: 6, 일: 0,
  };
  const thisWeekMatch = text.match(/이번\s*주?\s*([월화수목금토일])요일/);
  const nextWeekMatch = text.match(/다음\s*주\s*([월화수목금토일])요일/);
  if (thisWeekMatch || nextWeekMatch) {
    const m = (thisWeekMatch || nextWeekMatch)!;
    const target = dayMap[m[1]];
    const d = new Date(now);
    let diff = target - d.getDay();
    if (nextWeekMatch) diff += 7;
    else if (diff <= 0) diff += 7;
    d.setDate(d.getDate() + diff);
    return toISO(d);
  }

  const monthDay = text.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (monthDay) {
    const yr = now.getFullYear();
    return `${yr}-${monthDay[1].padStart(2, "0")}-${monthDay[2].padStart(2, "0")}`;
  }

  const fullDate = text.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (fullDate) {
    return `${fullDate[1]}-${fullDate[2].padStart(2, "0")}-${fullDate[3].padStart(2, "0")}`;
  }

  // M/D 또는 M-D (연도 생략)
  const shortDate = text.match(/\b(\d{1,2})[\/\-](\d{1,2})\b/);
  if (shortDate) {
    return `${now.getFullYear()}-${shortDate[1].padStart(2, "0")}-${shortDate[2].padStart(2, "0")}`;
  }

  if (/월말|이번달\s*말|말일/.test(text)) {
    const d = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    return toISO(d);
  }

  return null;
}

/** Slack 서명 검증 (HMAC-SHA256) */
function verifySlackSignature(req: NextRequest, rawBody: string): boolean {
  const secret = process.env.SLACK_SIGNING_SECRET;
  if (!secret) return true; // 환경변수 없으면 개발 모드로 통과

  const ts = req.headers.get("x-slack-request-timestamp") ?? "";
  const sig = req.headers.get("x-slack-signature") ?? "";

  // 리플레이 공격 방지 (5분 이내 요청만 허용)
  if (Math.abs(Date.now() / 1000 - Number(ts)) > 300) return false;

  const hmac = crypto.createHmac("sha256", secret);
  hmac.update(`v0:${ts}:${rawBody}`);
  const expected = `v0=${hmac.digest("hex")}`;

  try {
    return crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected));
  } catch {
    return false;
  }
}

/** Slack 메시지 전송 (스레드 답장 포함) */
async function sendSlack(channel: string, text: string, threadTs?: string) {
  const token = process.env.SLACK_BOT_TOKEN;
  if (!token) {
    console.log("[Slack] BOT_TOKEN 없음, 메시지 전송 생략:", text);
    return;
  }
  await fetch("https://slack.com/api/chat.postMessage", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ channel, text, thread_ts: threadTs }),
  });
}

// ─── 라우트 핸들러 ───

export async function POST(request: NextRequest) {
  const rawBody = await request.text();

  if (!verifySlackSignature(request, rawBody)) {
    return NextResponse.json({ error: "Invalid signature" }, { status: 401 });
  }

  const body = JSON.parse(rawBody);

  // Slack URL 검증 핸드셰이크
  if (body.type === "url_verification") {
    return NextResponse.json({ challenge: body.challenge });
  }

  if (body.type !== "event_callback") return NextResponse.json({ ok: true });

  const event = body.event;

  if (event.type !== "message") return NextResponse.json({ ok: true });

  // 메시지 삭제 이벤트 처리
  if (event.subtype === "message_deleted") {
    const deletedId = `slack-${event.deleted_ts}`;
    // confirmedTasks에서 제거
    const idx = confirmedTasks.findIndex((t) => t.id === deletedId);
    if (idx !== -1) confirmedTasks.splice(idx, 1);
    // awaitingDeadline에서도 제거
    awaitingDeadline.delete(event.deleted_ts);
    // 프론트에 삭제 신호 전달
    deletedTaskIds.push(deletedId);
    return NextResponse.json({ ok: true });
  }

  // 봇 메시지 / 기타 서브타입 무시
  if (event.bot_id || event.subtype) return NextResponse.json({ ok: true });

  const text: string = event.text || "";
  const channel: string = event.channel;
  const ts: string = event.ts;
  const threadTs: string | undefined = event.thread_ts;

  // ── Case 1: 마감기한 되묻기에 대한 스레드 답장 ──
  if (threadTs && awaitingDeadline.has(threadTs)) {
    const pending = awaitingDeadline.get(threadTs)!;
    const deadline = parseDeadline(text);

    if (deadline) {
      awaitingDeadline.delete(threadTs);
      const task: SlackTask = { ...pending, deadline };
      confirmedTasks.push(task);
      if (confirmedTasks.length > 100) confirmedTasks.shift();

      await sendSlack(
        channel,
        `✅ *${task.to}*님 업무가 대시보드에 등록됐습니다!\n• 마감: ${deadline}\n• 내용: ${task.message}`,
        threadTs
      );
    } else {
      // 날짜 인식 실패 → 다시 물어봄
      await sendSlack(
        channel,
        `날짜를 인식하지 못했어요 :sweat_smile:\n다음 형식으로 알려주세요:\n> "이번주 금요일" · "3월 28일" · "2026-03-31" · "내일"`,
        threadTs
      );
    }
    return NextResponse.json({ ok: true });
  }

  // ── Case 2: 신규 메시지 - @멘션이 있는 메시지만 처리 ──
  const mentioned = extractMentionedUser(text);
  if (!mentioned) return NextResponse.json({ ok: true });

  const fromName = USER_MAP[event.user] || event.user;
  const channelName = CHANNEL_MAP[channel] || channel;
  const cleanMessage = cleanSlackText(text);
  if (!cleanMessage) return NextResponse.json({ ok: true });

  const taskBase: PendingTask = {
    id: `slack-${ts}`,
    from: fromName,
    to: mentioned.name,
    message: cleanMessage,
    channel: channelName,
    timestamp: new Date(parseFloat(ts) * 1000).toLocaleString("ko-KR"),
    slackTs: ts,
  };

  const deadline = parseDeadline(text);

  if (!deadline) {
    // 마감기한 없음 → 스레드로 물어봄
    awaitingDeadline.set(ts, taskBase);
    await sendSlack(
      channel,
      `:wave: *${mentioned.name}*님 업무를 대시보드에 등록할게요.\n마감 기한을 확인할 수 없었어요. 언제까지 완료하면 될까요?`,
      ts
    );
  } else {
    // 마감기한 있음 → 바로 등록
    const task: SlackTask = { ...taskBase, deadline };
    confirmedTasks.push(task);
    if (confirmedTasks.length > 100) confirmedTasks.shift();

    await sendSlack(
      channel,
      `✅ *${mentioned.name}*님 업무가 대시보드에 등록됐습니다!\n• 마감: ${deadline}\n• 내용: ${cleanMessage}`,
      ts
    );
  }

  return NextResponse.json({ ok: true });
}

/** GET: 대시보드 폴링 — 신규 태스크 + 삭제 ID 반환 후 비움 */
export async function GET() {
  const tasks = [...confirmedTasks];
  const deletedIds = [...deletedTaskIds];
  confirmedTasks.length = 0;
  deletedTaskIds.length = 0;
  return NextResponse.json({ tasks, deletedIds });
}
