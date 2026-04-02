import { NextRequest, NextResponse } from "next/server";
import crypto from "crypto";
import path from "path";
import { upsertTask, removeTask, patchTask, getAllTasks } from "../../../lib/github-db";
import { postMessage, downloadFile, extractPptxFiles, extractTextFiles } from "../../../lib/slack";

/**
 * 슬랙 웹훅 수신 엔드포인트 v2
 *
 * v1 대비 추가:
 * - 분류 후 스레드에 안내 메시지 자동 발송
 * - 스레드 파일 업로드 감지 → Agent 자동 실행 → 결과 회신
 * - 장표 제작 멀티턴 세션 관리
 * - 예산품의 NL 파싱 자동 트리거
 *
 * 환경변수:
 *   SLACK_SIGNING_SECRET, SLACK_BOT_TOKEN, ANTHROPIC_API_KEY, GITHUB_TOKEN
 *
 * Slack 앱 권한:
 *   chat:write, channels:history, groups:history, files:read
 *
 * 모니터링 채널:
 *   #section-전략추진실-all (C08NNP1D3A9)
 *   #section-전략추진실-창준님 (C0AJ265GP8W)
 *   #wg-전략추진실 (C09L1LBK1GD)
 *   #wg-사업성장팀x전략추진실 (C0AL18T5KU6)
 */

// ─── 워크스페이스 경로 ───
const WORKSPACE = path.resolve(process.cwd(), "..");
const PPT_TRANSLATE_INPUT = path.join(WORKSPACE, "ppt-translater", "input");
const PPT_MAKER_DIR = path.join(WORKSPACE, "ppt-maker");
const MEETING_INPUT = path.join(WORKSPACE, "meeting-notes", "input");

// ─── 인메모리 스토어 ───

interface ThreadInfo {
  taskId: string;
  from: string;
  topicFile?: string;
  /** 장표 제작 세션용 */
  sessionId?: string;
  sessionStep?: number;
}

const threadRegistry = new Map<string, ThreadInfo>();

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

// ─── Agent 실행 헬퍼 ───

async function runAgent(topicFile: string, params?: Record<string, string>): Promise<{ success: boolean; stdout: string; stderr: string; error?: string }> {
  const baseUrl = process.env.VERCEL_URL
    ? `https://${process.env.VERCEL_URL}`
    : "http://localhost:3000";

  try {
    const res = await fetch(`${baseUrl}/api/agents`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ topicFile, params }),
    });
    return await res.json();
  } catch (err) {
    return { success: false, stdout: "", stderr: "", error: String(err) };
  }
}

// ─── Agent별 스레드 안내 메시지 ───

const AGENT_PROMPTS: Record<string, { needsFile: boolean; prompt: string }> = {
  "fluid/ppt-translate.md": {
    needsFile: true,
    prompt: "📎 번역할 PPT 파일을 이 스레드에 올려주세요.\n업로드하면 자동으로 번역을 시작합니다.",
  },
  "fluid/ppt-create.md": {
    needsFile: false,
    prompt: "📝 장표 제작을 시작합니다.\n어떤 내용의 장표인가요? 트랜스크립트, 초안, 또는 맥락을 이 스레드에 올려주세요.",
  },
  "fluid/meeting-notes.md": {
    needsFile: true,
    prompt: "🎙️ 회의록 TXT 파일을 이 스레드에 올려주세요.\n업로드하면 자동으로 요약을 시작합니다.",
  },
  "regular/macro-update.md": {
    needsFile: false,
    prompt: "📊 Macro 업데이트를 시작합니다. KOSIS 데이터를 확인하고 10개 시트를 업데이트합니다.\n잠시 기다려주세요...",
  },
  "regular/placement-analysis.md": {
    needsFile: false,
    prompt: "📈 Placement Survey 분석을 시작합니다.\n분기를 알려주세요 (예: 26Q1)",
  },
};

// ─── 유틸 함수 ───

function extractKeywords(text: string): string[] {
  const stopWords = new Set(["이거", "이게", "그거", "저거", "이건", "그건", "해줘", "해주세요", "부탁", "감사", "좋아", "알겠", "네", "예", "응", "ㅇㅇ", "ㅇㅋ", "ㄱㄱ", "일단", "그냥", "좀", "한번", "다시", "혹시"]);
  const words = text.split(/[\s,·\/\-\|]+/).map((w) => w.trim()).filter((w) => w.length >= 2 && !stopWords.has(w));
  return [...new Set(words)].slice(0, 3);
}

function extractMentionedUser(text: string): { id: string; name: string } | null {
  const match = text.match(/<@(U[A-Z0-9]+)>/);
  if (!match) return null;
  return { id: match[1], name: USER_MAP[match[1]] || match[1] };
}

function cleanSlackText(text: string): string {
  return text
    .replace(/<@U[A-Z0-9]+>/g, "")
    .replace(/<#C[A-Z0-9]+\|([^>]+)>/g, "#$1")
    .replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, "$2")
    .replace(/<(https?:\/\/[^>]+)>/g, "$1")
    .trim();
}

function parseDeadline(text: string): string | null {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const toISO = (d: Date) => d.toISOString().split("T")[0];

  if (/오늘/.test(text)) return toISO(now);
  if (/내일/.test(text)) { const d = new Date(now); d.setDate(d.getDate() + 1); return toISO(d); }
  if (/모레/.test(text)) { const d = new Date(now); d.setDate(d.getDate() + 2); return toISO(d); }

  const dayMap: Record<string, number> = { 월: 1, 화: 2, 수: 3, 목: 4, 금: 5, 토: 6, 일: 0 };
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
  if (monthDay) return `${now.getFullYear()}-${monthDay[1].padStart(2, "0")}-${monthDay[2].padStart(2, "0")}`;

  const fullDate = text.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (fullDate) return `${fullDate[1]}-${fullDate[2].padStart(2, "0")}-${fullDate[3].padStart(2, "0")}`;

  const shortDate = text.match(/\b(\d{1,2})[\/\-](\d{1,2})\b/);
  if (shortDate) return `${now.getFullYear()}-${shortDate[1].padStart(2, "0")}-${shortDate[2].padStart(2, "0")}`;

  if (/월말|이번달\s*말|말일/.test(text)) {
    const d = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    return toISO(d);
  }

  return null;
}

function verifySlackSignature(req: NextRequest, rawBody: string): boolean {
  const secret = process.env.SLACK_SIGNING_SECRET;
  if (!secret) return true;

  const ts = req.headers.get("x-slack-request-timestamp") ?? "";
  const sig = req.headers.get("x-slack-signature") ?? "";

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

// ─── topicFile → agent topicFile 매핑 (classify 결과 → agents API 키) ───
function classifyTopicToAgentTopic(classifyTopic: string): string | null {
  const map: Record<string, string | undefined> = {
    "regular/macro-update.md": "regular/macro-update",
    "regular/placement-analysis.md": "regular/placement-analysis",
    "fluid/ppt-translate.md": "fluid/ppt-translate",
    "fluid/ppt-create.md": "fluid/ppt-create",
    "fluid/meeting-notes.md": "fluid/meeting-notes",
    "fluid/budget-draft.md": "fluid/budget-draft",
  };
  // Handle both with and without .md
  const key = classifyTopic.endsWith(".md") ? classifyTopic : `${classifyTopic}.md`;
  if (key in map) return map[key] ?? null;
  // Fallback: strip .md
  return classifyTopic.replace(/\.md$/, "") || null;
}

// ─── 스레드 파일 처리 + Agent 실행 ───

async function handleThreadFile(
  channel: string,
  threadTs: string,
  info: ThreadInfo,
  files: { name: string; url_private: string; mimetype: string; filetype: string }[]
) {
  const agentTopic = info.topicFile ? classifyTopicToAgentTopic(info.topicFile) : null;

  // 번역 Agent: PPTX 파일 다운로드 → 번역 실행
  if (agentTopic === "fluid/ppt-translate") {
    const pptxFiles = extractPptxFiles(files);
    if (pptxFiles.length === 0) {
      await postMessage(channel, "⚠️ PPTX 파일이 아닙니다. .pptx 파일을 올려주세요.", threadTs);
      return;
    }

    const file = pptxFiles[0];
    await postMessage(channel, `📥 ${file.name} 다운로드 중...`, threadTs);

    const localPath = await downloadFile(file.url, file.name, PPT_TRANSLATE_INPUT);
    if (!localPath) {
      await postMessage(channel, "❌ 파일 다운로드 실패. 다시 시도해주세요.", threadTs);
      return;
    }

    await postMessage(channel, `🔄 번역 시작: ${file.name}\n잠시 기다려주세요 (3~5분 소요)`, threadTs);
    await patchTask(info.taskId, { status: "in_progress" });

    const result = await runAgent("fluid/ppt-translate", { input: `input/${file.name}` });

    if (result.success) {
      const outputName = file.name.replace(".pptx", "_en.pptx");
      await postMessage(
        channel,
        `✅ 번역 완료!\n📁 결과 파일: ppt-translate/output/${outputName}\n\n검수 후 사용해주세요.`,
        threadTs
      );
      await patchTask(info.taskId, { status: "done" });
    } else {
      await postMessage(
        channel,
        `❌ 번역 실패\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``,
        threadTs
      );
    }
    return;
  }

  // 회의록 Agent: TXT 파일 다운로드 → 요약 실행
  if (agentTopic === "fluid/meeting-notes") {
    const txtFiles = extractTextFiles(files);
    if (txtFiles.length === 0) {
      await postMessage(channel, "⚠️ TXT 파일이 아닙니다. .txt 파일을 올려주세요.", threadTs);
      return;
    }

    const file = txtFiles[0];
    const localPath = await downloadFile(file.url, file.name, MEETING_INPUT);
    if (!localPath) {
      await postMessage(channel, "❌ 파일 다운로드 실패.", threadTs);
      return;
    }

    await postMessage(channel, `🔄 회의록 요약 시작: ${file.name}`, threadTs);
    await patchTask(info.taskId, { status: "in_progress" });

    const result = await runAgent("fluid/meeting-notes", { input: `input/${file.name}` });

    if (result.success) {
      await postMessage(channel, `✅ 회의록 요약 완료!\n${result.stdout.slice(0, 500)}`, threadTs);
      await patchTask(info.taskId, { status: "done" });
    } else {
      await postMessage(channel, `❌ 요약 실패\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``, threadTs);
    }
    return;
  }
}

// ─── 스레드 텍스트 처리 (품의, Placement, 장표 제작) ───

async function handleThreadText(
  channel: string,
  threadTs: string,
  info: ThreadInfo,
  text: string
) {
  const agentTopic = info.topicFile ? classifyTopicToAgentTopic(info.topicFile) : null;

  // Placement: 분기 입력 → 실행
  if (agentTopic === "regular/placement-analysis") {
    const quarterMatch = text.match(/(\d{2})[Qq](\d)/);
    if (quarterMatch) {
      const quarter = `${quarterMatch[1]}Q${quarterMatch[2]}`;
      await postMessage(channel, `📈 Placement ${quarter} 분석을 시작합니다. 잠시 기다려주세요...`, threadTs);
      await patchTask(info.taskId, { status: "in_progress" });

      const result = await runAgent("regular/placement-analysis", { quarter });

      if (result.success) {
        await postMessage(channel, `✅ Placement ${quarter} 분석 완료!\n${result.stdout.slice(0, 500)}`, threadTs);
        await patchTask(info.taskId, { status: "done" });
      } else {
        await postMessage(channel, `❌ 분석 실패\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``, threadTs);
      }
      return;
    }
  }

  // 예산품의: NL 메시지 → auto.js
  if (agentTopic === "fluid/budget-draft") {
    await postMessage(channel, `📋 품의 처리를 시작합니다: "${text.slice(0, 100)}"`, threadTs);
    await patchTask(info.taskId, { status: "in_progress" });

    const result = await runAgent("fluid/budget-draft", { message: text });

    if (result.success) {
      await postMessage(channel, `✅ 품의 자동 입력 완료 (dry-run)\n\`\`\`${result.stdout.slice(0, 500)}\`\`\`\n\n확인 후 실제 제출하려면 대시보드에서 진행해주세요.`, threadTs);
      await patchTask(info.taskId, { status: "done" });
    } else {
      await postMessage(channel, `❌ 품의 처리 실패\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``, threadTs);
    }
    return;
  }

  // 장표 제작: 멀티턴 세션
  if (agentTopic === "fluid/ppt-create") {
    const sessionId = info.sessionId || `ppt-${Date.now()}`;
    const step = info.sessionStep || 1;

    // 세션 시작 or 진행
    const result = await runAgent("fluid/ppt-create", {
      session: sessionId,
      step: String(step),
      input: text,
    });

    if (result.success) {
      try {
        const parsed = JSON.parse(result.stdout);
        // 다음 스텝 정보 업데이트
        info.sessionId = sessionId;
        info.sessionStep = parsed.nextStep || step + 1;
        threadRegistry.set(threadTs, info);

        // 결과 메시지 전송
        let replyText = parsed.message || "처리 완료";
        if (parsed.prompt) replyText += `\n\n${parsed.prompt}`;
        if (parsed.data?.headMessages) {
          replyText += "\n\n" + parsed.data.headMessages
            .map((h: { slide: number; message: string }, i: number) => `${i + 1}. ${h.message}`)
            .join("\n");
        }
        await postMessage(channel, replyText, threadTs);

        // 최종 단계면 완료 처리
        if (parsed.step >= 7 && parsed.data?.outputPath) {
          await postMessage(channel, `📁 PPTX 저장 완료: ${parsed.data.outputPath}`, threadTs);
          await patchTask(info.taskId, { status: "done" });
        }
      } catch {
        // JSON 파싱 실패 시 raw output 전송
        await postMessage(channel, result.stdout.slice(0, 500), threadTs);
      }
    } else {
      await postMessage(channel, `❌ 장표 제작 오류\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``, threadTs);
    }
    return;
  }
}

// ─── 즉시 실행 Agent 처리 (파라미터 불필요) ───

async function handleImmediateAgent(
  channel: string,
  threadTs: string,
  info: ThreadInfo,
  message: string
) {
  const agentTopic = info.topicFile ? classifyTopicToAgentTopic(info.topicFile) : null;

  // 매크로: 바로 실행
  if (agentTopic === "regular/macro-update") {
    await patchTask(info.taskId, { status: "in_progress" });
    const result = await runAgent("regular/macro-update");

    if (result.success) {
      await postMessage(channel, `✅ Macro 업데이트 완료!\n${result.stdout.slice(0, 500)}`, threadTs);
      await patchTask(info.taskId, { status: "done" });
    } else {
      await postMessage(channel, `❌ Macro 업데이트 실패\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``, threadTs);
    }
    return true;
  }

  // 예산품의: 메시지 자체가 NL 입력
  if (agentTopic === "fluid/budget-draft") {
    await postMessage(channel, `📋 품의 처리를 시작합니다: "${message.slice(0, 100)}"`, threadTs);
    await patchTask(info.taskId, { status: "in_progress" });

    const result = await runAgent("fluid/budget-draft", { message });

    if (result.success) {
      await postMessage(channel, `✅ 품의 자동 입력 완료 (dry-run)\n\`\`\`${result.stdout.slice(0, 500)}\`\`\`\n\n확인 후 실제 제출하려면 대시보드에서 진행해주세요.`, threadTs);
      await patchTask(info.taskId, { status: "done" });
    } else {
      await postMessage(channel, `❌ 품의 처리 실패\n\`\`\`${(result.stderr || result.error || "").slice(0, 300)}\`\`\``, threadTs);
    }
    return true;
  }

  return false;
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

  // ── 메시지 삭제 → GitHub에서 제거 ──
  if (event.subtype === "message_deleted") {
    const deletedId = `slack-${event.deleted_ts}`;
    threadRegistry.delete(event.deleted_ts);
    await removeTask(deletedId).catch(() => {});
    return NextResponse.json({ ok: true });
  }

  // 봇 메시지 무시 (단, file_share는 처리)
  if (event.bot_id) return NextResponse.json({ ok: true });
  if (event.subtype && event.subtype !== "file_share") return NextResponse.json({ ok: true });

  const text: string = event.text || "";
  const channel: string = event.channel;
  const ts: string = event.ts;
  const threadTs: string | undefined = event.thread_ts;
  const files = event.files || [];

  // ══════════════════════════════════════════════════════
  // Case 1: 스레드 후속 메시지 (파일 또는 텍스트)
  // ══════════════════════════════════════════════════════
  if (threadTs && threadRegistry.has(threadTs)) {
    const info = threadRegistry.get(threadTs)!;
    const fromName = USER_MAP[event.user] || event.user;
    const cleanedText = cleanSlackText(text);

    // 파일이 있으면 → Agent 파일 처리
    if (files.length > 0) {
      handleThreadFile(channel, threadTs, info, files).catch(console.error);
      return NextResponse.json({ ok: true });
    }

    // 텍스트가 있으면 → Agent 텍스트 처리 or 기존 보강
    if (cleanedText) {
      // Agent가 할당된 경우 텍스트 기반 처리 시도
      if (info.topicFile) {
        handleThreadText(channel, threadTs, info, cleanedText).catch(console.error);
        return NextResponse.json({ ok: true });
      }

      // Agent 미할당: 기존 로직 (내용/마감기한 보강)
      const all = await getAllTasks();
      const existing = all.find((t) => t.id === info.taskId);
      if (!existing) return NextResponse.json({ ok: true });

      if (info.from === fromName) {
        const newDeadline = parseDeadline(text);
        await patchTask(info.taskId, {
          message: (existing.message + " / " + cleanedText).slice(0, 300),
          ...(existing.deadline === "미정" && newDeadline ? { deadline: newDeadline } : {}),
        });
      } else {
        const keywords = extractKeywords(cleanedText);
        if (keywords.length > 0) {
          const currentNotes: string[] = (existing as { notes?: string[] }).notes || [];
          await patchTask(info.taskId, { notes: [...currentNotes, ...keywords].slice(-10) });
        }
      }
    }

    return NextResponse.json({ ok: true });
  }

  // ══════════════════════════════════════════════════════
  // Case 2: 신규 메시지 — @멘션이 있는 메시지만 처리
  // ══════════════════════════════════════════════════════
  const mentioned = extractMentionedUser(text);
  if (!mentioned) return NextResponse.json({ ok: true });

  const fromName = USER_MAP[event.user] || event.user;
  const channelName = CHANNEL_MAP[channel] || channel;
  const cleanMessage = cleanSlackText(text);
  if (!cleanMessage) return NextResponse.json({ ok: true });

  // ── 사담 필터링: Claude API로 업무 여부 판별 ──
  try {
    const baseUrl = process.env.VERCEL_URL
      ? `https://${process.env.VERCEL_URL}`
      : "http://localhost:3000";
    const filterRes = await fetch(`${baseUrl}/api/classify`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: cleanMessage, checkWorkOnly: true }),
    });
    const filterData = await filterRes.json();
    if (filterData.isWork === false) {
      console.log(`[slack-webhook] 사담 필터링됨: "${cleanMessage.slice(0, 50)}"`);
      return NextResponse.json({ ok: true });
    }
  } catch {
    // 필터링 실패 시 안전하게 통과 (업무로 간주)
  }

  const taskId = `slack-${ts}`;
  const task = {
    id: taskId,
    from: fromName,
    to: mentioned.name,
    message: cleanMessage,
    channel: channelName,
    timestamp: new Date(parseFloat(ts) * 1000).toLocaleString("ko-KR", { timeZone: "Asia/Seoul" }),
    deadline: parseDeadline(text) ?? "미정",
    status: "pending" as const,
    slackTs: ts,
  };

  // GitHub에 저장 + 스레드 추적
  await upsertTask(task);
  threadRegistry.set(ts, { taskId, from: fromName });

  // 비동기: classify → 스레드 안내 메시지 → 즉시 실행 Agent 처리
  const baseUrl = process.env.VERCEL_URL
    ? `https://${process.env.VERCEL_URL}`
    : "http://localhost:3000";

  fetch(`${baseUrl}/api/classify`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ message: cleanMessage }),
  })
    .then((r) => r.json())
    .then(async (classified) => {
      if (!classified.category || classified.category === "미분류") return;

      // GitHub 업데이트
      await patchTask(taskId, {
        category: classified.category,
        topicFile: classified.topicFile,
        autoLevel: classified.autoLevel,
        guide: classified.guide,
        steps: classified.steps,
      });

      // 스레드 레지스트리에 topicFile 저장
      const info = threadRegistry.get(ts);
      if (info) {
        info.topicFile = classified.topicFile;
        threadRegistry.set(ts, info);
      }

      // auto Agent만 처리
      if (classified.autoLevel !== "auto") return;

      // 안내 메시지 전송
      const prompt = AGENT_PROMPTS[classified.topicFile];
      if (prompt) {
        await postMessage(channel, prompt.prompt, ts);
      }

      // 파일 불필요한 Agent → 즉시 실행
      if (prompt && !prompt.needsFile) {
        const updatedInfo = threadRegistry.get(ts);
        if (updatedInfo) {
          await handleImmediateAgent(channel, ts, updatedInfo, cleanMessage);
        }
      }
    })
    .catch(console.error);

  return NextResponse.json({ ok: true });
}

/** GET: 상태 확인용 */
export async function GET() {
  return NextResponse.json({ ok: true, threads: threadRegistry.size });
}
