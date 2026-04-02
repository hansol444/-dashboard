import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import PptxGenJS from "pptxgenjs";

/**
 * 장표 제작 API — Vercel 완전 호환
 *
 * 세션 상태를 클라이언트가 관리하므로 서버 파일시스템 불필요.
 * POST body: { step, userInput, session }
 * - session: 클라이언트가 보관하는 세션 객체 (이전 응답의 session 필드)
 * - step 7 응답에 pptxBase64 포함 → 클라이언트에서 다운로드
 */

const MODEL = "claude-sonnet-4-5-20250929";

// ─── Claude 호출 헬퍼 ──────────────────────────────────────────────────────

function getClient() {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY가 설정되지 않았습니다.");
  return new Anthropic({ apiKey });
}

async function callClaude(system: string, userMsg: string): Promise<string> {
  const client = getClient();
  const resp = await client.messages.create({
    model: MODEL,
    max_tokens: 4096,
    system,
    messages: [{ role: "user", content: userMsg }],
  });
  return resp.content[0].type === "text" ? resp.content[0].text : "";
}

function parseJson(raw: string): Record<string, unknown> {
  // ```json ... ``` 또는 { ... } 추출
  let text = raw;
  if (text.includes("```json")) {
    text = text.split("```json")[1].split("```")[0];
  } else if (text.includes("```")) {
    text = text.split("```")[1].split("```")[0];
  }
  const match = text.match(/\{[\s\S]*\}/);
  if (match) return JSON.parse(match[0]);
  throw new Error("JSON을 찾을 수 없습니다");
}

async function callClaudeJson(system: string, userMsg: string): Promise<Record<string, unknown>> {
  const raw = await callClaude(system, userMsg);
  return parseJson(raw);
}

// ─── 세션 타입 ──────────────────────────────────────────────────────────────

interface Session {
  id: string;
  step: number;
  context: string;
  audience: string;
  purpose: string;
  headMessages: Array<{
    slideNumber: number;
    headMessage: string;
    subPoints: string[];
    role: string;
    checklist?: Record<string, boolean>;
  }>;
  layouts: Array<{
    slideNumber: number;
    layout: string;
    reason: string;
    contentStructure?: { type: string; details: string };
  }>;
}

function newSession(id?: string): Session {
  return {
    id: id || `ppt-${Date.now()}`,
    step: 0,
    context: "",
    audience: "",
    purpose: "",
    headMessages: [],
    layouts: [],
  };
}

// ─── 응답 포맷 ──────────────────────────────────────────────────────────────

interface StepResponse {
  step: number;
  status: string;
  message: string;
  data: Record<string, unknown>;
  nextStep: number;
  needsInput: boolean;
  prompt: string;
  session: Session;
  pptxBase64?: string;
}

function makeResponse(
  step: number, status: string, message: string,
  data: Record<string, unknown>, nextStep: number,
  needsInput: boolean, prompt: string, session: Session,
  pptxBase64?: string,
): StepResponse {
  return { step, status, message, data, nextStep, needsInput, prompt, session, pptxBase64 };
}

// ─── Step 1: 입력 분석 ──────────────────────────────────────────────────────

async function step1(session: Session, userInput: string): Promise<StepResponse> {
  session.context = userInput;

  const result = await callClaudeJson(
    "당신은 프레젠테이션 제작 전문가입니다. 사용자가 제공한 텍스트의 유형을 분석하세요.\n" +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"inputType":"transcript|draft|conversation|document|bullet_points",' +
    '"summary":"핵심 내용 요약 (3문장 이내)",' +
    '"keyTopics":["주제1","주제2"],' +
    '"estimatedSlides":숫자,' +
    '"language":"ko|en"}',
    `다음 텍스트를 분석하세요:\n\n${userInput.slice(0, 3000)}`
  );

  session.step = 1;

  return makeResponse(1, "done", "입력을 분석했습니다.", result, 2, true,
    "입력이 분석되었습니다. 다음 질문에 답변해 주세요:\n" +
    "1) 누구 입장에서 발표하나요? (대상 청중)\n" +
    "2) 어느 깊이까지 다루나요? (개요/상세)\n" +
    "3) 현재 상태는? (초안/최종)",
    session);
}

// ─── Step 2: 명확화 질문 ────────────────────────────────────────────────────

async function step2(session: Session, userInput: string): Promise<StepResponse> {
  const result = await callClaudeJson(
    "당신은 프레젠테이션 기획자입니다. " +
    "사용자의 원본 컨텍스트와 답변을 바탕으로 청중, 목적, 깊이를 정리하세요.\n" +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"audience":"청중 설명","purpose":"발표 목적","depth":"개요|상세|전략 브리핑",' +
    '"suggestedSlideCount":숫자,"slideTopics":["주제1","주제2"]}',
    `원본 컨텍스트:\n${session.context.slice(0, 2000)}\n\n사용자 답변:\n${userInput}`
  );

  session.audience = (result.audience as string) || "";
  session.purpose = (result.purpose as string) || "";
  session.step = 2;

  return makeResponse(2, "done", "명확화가 완료되었습니다.", result, 3, false,
    "다음 단계에서 헤드메시지 초안을 생성합니다.", session);
}

// ─── Step 3: 헤드메시지 초안 ────────────────────────────────────────────────

async function step3(session: Session): Promise<StepResponse> {
  const result = await callClaudeJson(
    "당신은 프레젠테이션 헤드메시지 전문가입니다.\n" +
    "주어진 컨텍스트, 청중, 목적을 기반으로 슬라이드별 헤드메시지를 작성하세요.\n\n" +
    "9-체크리스트 검증 기준:\n" +
    "1. 한 문장으로 핵심 메시지 전달\n" +
    "2. So What 테스트 통과 (행동/판단 유도)\n" +
    "3. 수치/근거 포함 여부\n" +
    "4. 청중 관점에서 의미 있는 내용\n" +
    "5. 이전 슬라이드와 논리적 연결\n" +
    "6. 중복 없음\n" +
    "7. 추상적 표현 배제\n" +
    "8. 15단어 이내 간결성\n" +
    "9. 전체 스토리라인에서 역할 명확\n\n" +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"slides":[{"slideNumber":1,"headMessage":"...","subPoints":["..."],' +
    '"checklist":{"soWhat":true,"hasEvidence":true},"role":"도입|본론|결론"}]}',
    `컨텍스트:\n${session.context.slice(0, 2000)}\n\n청중: ${session.audience}\n목적: ${session.purpose}`
  );

  const slides = (result.slides as Session["headMessages"]) || [];
  session.headMessages = slides;
  session.step = 3;

  return makeResponse(3, "done", "헤드메시지 초안을 생성했습니다.",
    { headMessages: slides }, 4, true,
    "초안을 검토하고 수정사항이 있으면 알려주세요. 없으면 '확정'이라고 해주세요.",
    session);
}

// ─── Step 4: 사용자 확인 ────────────────────────────────────────────────────

async function step4(session: Session, userInput: string): Promise<StepResponse> {
  if (userInput.includes("확정")) {
    session.step = 4;
    return makeResponse(4, "done", "헤드메시지가 확정되었습니다.",
      { confirmed: true, slides: session.headMessages }, 5, false,
      "다음 단계에서 레이아웃을 자동 선택합니다.", session);
  }

  // 수정 요청 처리
  const result = await callClaudeJson(
    "당신은 프레젠테이션 헤드메시지 수정 전문가입니다.\n" +
    "현재 헤드메시지 초안과 사용자의 수정 요청을 반영하여 수정된 버전을 반환하세요.\n" +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"slides":[{"slideNumber":1,"headMessage":"...","subPoints":["..."],"role":"도입|본론|결론"}]}',
    `현재 초안:\n${JSON.stringify(session.headMessages)}\n\n수정 요청:\n${userInput}`
  );

  const slides = (result.slides as Session["headMessages"]) || [];
  session.headMessages = slides;
  session.step = 3;

  return makeResponse(4, "revised", "수정된 헤드메시지를 반영했습니다.",
    { headMessages: slides }, 4, true,
    "수정된 초안을 확인하세요. 추가 수정이 필요하면 알려주세요. 없으면 '확정'이라고 해주세요.",
    session);
}

// ─── Step 5: 레이아웃 선택 ──────────────────────────────────────────────────

async function step5(session: Session): Promise<StepResponse> {
  const result = await callClaudeJson(
    "당신은 프레젠테이션 레이아웃 전문가입니다.\n" +
    "각 슬라이드의 헤드메시지와 내용을 기반으로 최적의 레이아웃 유형을 선택하세요.\n\n" +
    "레이아웃 유형:\n" +
    '- "텍스트형": 텍스트 중심 슬라이드\n' +
    '- "표형": 표/비교 슬라이드\n' +
    '- "차트형": 차트/그래프 슬라이드\n' +
    '- "다이어그램형": 프로세스/플로우 다이어그램\n' +
    '- "타임라인형": 타임라인 슬라이드\n' +
    '- "하이브리드형": 복합 레이아웃\n\n' +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"layouts":[{"slideNumber":1,"layout":"텍스트형","reason":"이유",' +
    '"contentStructure":{"type":"bullets|table|chart|diagram|timeline|mixed","details":"..."}}]}',
    `슬라이드 목록:\n${JSON.stringify(session.headMessages)}`
  );

  const layouts = (result.layouts as Session["layouts"]) || [];
  session.layouts = layouts;
  session.step = 5;

  return makeResponse(5, "done", "레이아웃을 자동 선택했습니다.",
    { layouts }, 6, true,
    "레이아웃을 확인하세요. 변경이 필요하면 알려주세요. 없으면 '확정'이라고 해주세요.",
    session);
}

// ─── Step 6: 레이아웃 확인 ──────────────────────────────────────────────────

async function step6(session: Session, userInput: string): Promise<StepResponse> {
  if (userInput.includes("확정")) {
    session.step = 6;
    return makeResponse(6, "done", "레이아웃이 확정되었습니다.",
      { confirmed: true, layouts: session.layouts }, 7, false,
      "다음 단계에서 PPTX를 생성합니다.", session);
  }

  const result = await callClaudeJson(
    "당신은 프레젠테이션 레이아웃 수정 전문가입니다.\n" +
    "현재 레이아웃 설정과 사용자의 수정 요청을 반영하여 수정된 버전을 반환하세요.\n" +
    "레이아웃 유형: 텍스트형, 표형, 차트형, 다이어그램형, 타임라인형, 하이브리드형\n" +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"layouts":[{"slideNumber":1,"layout":"텍스트형","reason":"..."}]}',
    `현재 레이아웃:\n${JSON.stringify(session.layouts)}\n\n수정 요청:\n${userInput}`
  );

  const layouts = (result.layouts as Session["layouts"]) || [];
  session.layouts = layouts;
  session.step = 5;

  return makeResponse(6, "revised", "수정된 레이아웃을 반영했습니다.",
    { layouts }, 6, true,
    "수정된 레이아웃을 확인하세요. 추가 수정이 필요하면 알려주세요. 없으면 '확정'이라고 해주세요.",
    session);
}

// ─── Step 7: PPTX 생성 ─────────────────────────────────────────────────────

interface SlideContent {
  slideNumber: number;
  title: string;
  layout: string;
  content: Record<string, unknown>;
}

async function generateSlideContent(session: Session): Promise<SlideContent[]> {
  const slidesInfo = session.headMessages.map((hm, i) => ({
    slideNumber: hm.slideNumber,
    headMessage: hm.headMessage,
    subPoints: hm.subPoints,
    layout: session.layouts[i]?.layout || "텍스트형",
  }));

  const result = await callClaudeJson(
    "당신은 프레젠테이션 콘텐츠 작성 전문가입니다.\n" +
    "각 슬라이드의 헤드메시지, 레이아웃 유형, 원본 컨텍스트를 기반으로 " +
    "슬라이드에 들어갈 구체적인 콘텐츠를 생성하세요.\n\n" +
    "레이아웃별 콘텐츠 형식:\n" +
    "- 텍스트형: bullets 배열 (각 항목은 문자열)\n" +
    '- 표형: table 객체 {headers: [...], rows: [[...], [...]]}\n' +
    '- 차트형: chart 객체 {chartType: "bar|line|pie", labels: [...], values: [...], title: "..."}\n' +
    '- 다이어그램형: steps 배열 [{label: "...", description: "..."}]\n' +
    '- 타임라인형: events 배열 [{date: "...", title: "...", description: "..."}]\n' +
    '- 하이브리드형: {text: [...], chart: {chartType: "...", labels: [...], values: [...]}}\n\n' +
    "반드시 JSON만 반환하세요.\n" +
    'JSON: {"slides":[{"slideNumber":1,"title":"...","layout":"텍스트형","content":{...}}]}',
    `컨텍스트:\n${session.context.slice(0, 2000)}\n\n` +
    `청중: ${session.audience}\n목적: ${session.purpose}\n\n` +
    `슬라이드 구성:\n${JSON.stringify(slidesInfo)}`
  );

  return (result.slides as SlideContent[]) || [];
}

// pptxgenjs 색상/설정
const C = {
  navy: "1B365D",
  accent: "00B388",
  white: "FFFFFF",
  dark: "333333",
  gray: "F2F2F2",
  darkGray: "718096",
};
const FONT = "Malgun Gothic";

function buildTextSlide(slide: PptxGenJS.Slide, data: SlideContent) {
  const content = data.content || {};
  const bullets: string[] = Array.isArray(content)
    ? content
    : (content.bullets as string[]) || [];

  if (bullets.length === 0) return;

  const bodyTexts = bullets.map((b) => ({
    text: typeof b === "object" ? JSON.stringify(b) : String(b),
    options: {
      fontSize: 14,
      color: C.dark,
      fontFace: FONT,
      bullet: { code: "2022" as const },
      breakLine: true as const,
      paraSpaceAfter: 8,
    },
  }));

  slide.addText(bodyTexts, { x: 0.8, y: 1.4, w: 11.7, h: 5.0, valign: "top" });
}

function buildTableSlide(slide: PptxGenJS.Slide, data: SlideContent) {
  const content = data.content || {};
  const tbl = (content.headers ? content : content.table) as {
    headers?: string[];
    rows?: string[][];
  } || {};
  const headers = tbl.headers || ["항목", "내용"];
  const rows = tbl.rows || [["데이터 없음", "-"]];

  const tableRows: PptxGenJS.TableRow[] = [];

  // 헤더 행
  tableRows.push(
    headers.map((h) => ({
      text: String(h),
      options: {
        bold: true,
        fontSize: 11,
        color: C.white,
        fill: { color: C.navy },
        fontFace: FONT,
        align: "center" as const,
      },
    }))
  );

  // 데이터 행
  rows.forEach((row, i) => {
    tableRows.push(
      headers.map((_, j) => ({
        text: String(row[j] ?? ""),
        options: {
          fontSize: 10,
          color: C.dark,
          fontFace: FONT,
          fill: i % 2 === 1 ? { color: C.gray } : undefined,
        },
      }))
    );
  });

  slide.addTable(tableRows, {
    x: 0.8, y: 1.4, w: 11.7,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    colW: headers.map(() => 11.7 / headers.length),
    autoPage: true,
  });
}

function buildChartSlide(slide: PptxGenJS.Slide, data: SlideContent) {
  const content = data.content || {};
  const chart = (content.labels ? content : content.chart) as {
    chartType?: string;
    labels?: string[];
    values?: number[];
    title?: string;
  } || {};

  const labels = chart.labels || ["A", "B", "C"];
  const values = chart.values || [30, 50, 20];
  const title = chart.title || "";
  const maxVal = Math.max(...values, 1);
  const barMaxW = 8.0;

  if (title) {
    slide.addText(title, {
      x: 0.8, y: 1.4, w: 11.7, h: 0.4,
      fontSize: 14, bold: true, color: C.navy, fontFace: FONT, align: "center",
    });
  }

  const startY = title ? 2.0 : 1.6;
  const barH = Math.min(0.5, 4.5 / Math.max(labels.length, 1) - 0.1);

  labels.forEach((label, i) => {
    const y = startY + (barH + 0.15) * i;
    // 라벨
    slide.addText(String(label), {
      x: 0.8, y, w: 1.8, h: barH,
      fontSize: 10, color: C.dark, fontFace: FONT, align: "right", valign: "middle",
    });
    // 막대
    const w = Math.max(0.2, barMaxW * (values[i] / maxVal));
    slide.addShape("rect" as PptxGenJS.ShapeType, {
      x: 2.8, y, w, h: barH,
      fill: { color: C.accent },
      line: { type: "none" },
    });
    // 값
    slide.addText(String(values[i]), {
      x: 2.8 + w + 0.1, y, w: 0.8, h: barH,
      fontSize: 10, bold: true, color: C.navy, fontFace: FONT, valign: "middle",
    });
  });
}

function buildDiagramSlide(slide: PptxGenJS.Slide, data: SlideContent) {
  const content = data.content || {};
  const steps: Array<{ label: string; description: string }> = Array.isArray(content)
    ? content
    : (content.steps as Array<{ label: string; description: string }>) || [];

  if (steps.length === 0) return;

  const n = steps.length;
  const stepW = Math.min(2.2, 11.0 / n - 0.2);
  const gap = 0.3;
  const totalW = stepW * n + gap * (n - 1);
  const startX = (13.333 - totalW) / 2;

  steps.forEach((s, i) => {
    const x = startX + (stepW + gap) * i;
    const info = typeof s === "string" ? { label: s, description: "" } : s;

    // 번호 원
    slide.addShape("ellipse" as PptxGenJS.ShapeType, {
      x: x + (stepW - 0.6) / 2, y: 1.4, w: 0.6, h: 0.6,
      fill: { color: C.accent }, line: { type: "none" },
    });
    slide.addText(String(i + 1), {
      x: x + (stepW - 0.6) / 2, y: 1.4, w: 0.6, h: 0.6,
      fontSize: 14, bold: true, color: C.white, fontFace: FONT, align: "center", valign: "middle",
    });

    // 라벨 박스
    slide.addShape("rect" as PptxGenJS.ShapeType, {
      x, y: 2.2, w: stepW, h: 1.8,
      fill: { color: C.gray },
      line: { color: C.accent, pt: 1.5 },
    });
    slide.addText(info.label || `단계 ${i + 1}`, {
      x, y: 2.2, w: stepW, h: 0.6,
      fontSize: 12, bold: true, color: C.navy, fontFace: FONT, align: "center", valign: "middle",
    });
    if (info.description) {
      slide.addText(info.description, {
        x: x + 0.1, y: 2.8, w: stepW - 0.2, h: 1.1,
        fontSize: 9, color: C.dark, fontFace: FONT, align: "center", valign: "top",
      });
    }

    // 화살표
    if (i < n - 1) {
      slide.addShape("rect" as PptxGenJS.ShapeType, {
        x: x + stepW + 0.05, y: 3.0, w: gap - 0.1, h: 0.05,
        fill: { color: C.accent }, line: { type: "none" },
      });
    }
  });
}

function buildTimelineSlide(slide: PptxGenJS.Slide, data: SlideContent) {
  const content = data.content || {};
  const events: Array<{ date: string; title: string; description: string }> =
    Array.isArray(content)
      ? content
      : (content.events as Array<{ date: string; title: string; description: string }>) || [];

  if (events.length === 0) return;

  const lineY = 3.0;
  // 수평 타임라인 선
  slide.addShape("rect" as PptxGenJS.ShapeType, {
    x: 1.0, y: lineY, w: 11.333, h: 0.06,
    fill: { color: C.navy }, line: { type: "none" },
  });

  const itemW = 11.333 / events.length;

  events.forEach((ev, i) => {
    const info = typeof ev === "string" ? { date: "", title: ev, description: "" } : ev;
    const xCenter = 1.0 + itemW * i + itemW / 2;

    // 원형 마커
    slide.addShape("ellipse" as PptxGenJS.ShapeType, {
      x: xCenter - 0.15, y: lineY - 0.12, w: 0.3, h: 0.3,
      fill: { color: C.accent }, line: { type: "none" },
    });

    // 날짜 (위)
    slide.addText(info.date || "", {
      x: xCenter - 0.8, y: lineY - 0.8, w: 1.6, h: 0.5,
      fontSize: 10, bold: true, color: C.accent, fontFace: FONT, align: "center",
    });

    // 제목 (아래)
    slide.addText(info.title || "", {
      x: xCenter - 0.9, y: lineY + 0.35, w: 1.8, h: 0.4,
      fontSize: 11, bold: true, color: C.navy, fontFace: FONT, align: "center",
    });

    // 설명
    if (info.description) {
      slide.addText(info.description, {
        x: xCenter - 0.9, y: lineY + 0.75, w: 1.8, h: 0.8,
        fontSize: 9, color: C.dark, fontFace: FONT, align: "center", valign: "top",
      });
    }
  });
}

function buildHybridSlide(slide: PptxGenJS.Slide, data: SlideContent) {
  const content = data.content || {};
  const textItems: string[] = (content.text as string[]) || (content.bullets as string[]) || [];
  const chart = (content.chart as {
    labels?: string[];
    values?: number[];
  }) || {};

  // 왼쪽: 텍스트
  if (textItems.length > 0) {
    const bodyTexts = textItems.map((t) => ({
      text: typeof t === "object" ? JSON.stringify(t) : String(t),
      options: {
        fontSize: 12, color: C.dark, fontFace: FONT,
        bullet: { code: "2022" as const }, breakLine: true as const, paraSpaceAfter: 6,
      },
    }));
    slide.addText(bodyTexts, { x: 0.8, y: 1.4, w: 5.5, h: 5.0, valign: "top" });
  }

  // 오른쪽: 간단 차트
  const labels = chart.labels || [];
  const values = chart.values || [];
  if (labels.length > 0) {
    const maxVal = Math.max(...values, 1);
    const barMaxW = 3.5;
    const barH = Math.min(0.4, 4.5 / labels.length - 0.1);

    labels.forEach((label, i) => {
      const y = 1.6 + (barH + 0.15) * i;
      slide.addText(String(label), {
        x: 6.8, y, w: 1.5, h: barH,
        fontSize: 9, color: C.dark, fontFace: FONT, align: "right", valign: "middle",
      });
      const w = Math.max(0.2, barMaxW * ((values[i] || 0) / maxVal));
      slide.addShape("rect" as PptxGenJS.ShapeType, {
        x: 8.5, y, w, h: barH,
        fill: { color: C.accent }, line: { type: "none" },
      });
      slide.addText(String(values[i] || 0), {
        x: 8.5 + w + 0.05, y, w: 0.6, h: barH,
        fontSize: 9, bold: true, color: C.navy, fontFace: FONT, valign: "middle",
      });
    });
  }
}

const LAYOUT_BUILDERS: Record<string, (slide: PptxGenJS.Slide, data: SlideContent) => void> = {
  "텍스트형": buildTextSlide,
  "표형": buildTableSlide,
  "차트형": buildChartSlide,
  "다이어그램형": buildDiagramSlide,
  "타임라인형": buildTimelineSlide,
  "하이브리드형": buildHybridSlide,
};

async function step7(session: Session): Promise<StepResponse> {
  // 슬라이드 콘텐츠 생성 (Claude)
  const slidesContent = await generateSlideContent(session);

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "전략추진실 장표 제작 도우미";

  for (let i = 0; i < slidesContent.length; i++) {
    const sd = slidesContent[i];
    const slide = pptx.addSlide();

    slide.background = { color: C.white };

    const title = sd.title || session.headMessages[i]?.headMessage || `슬라이드 ${i + 1}`;
    const layoutType = sd.layout || session.layouts[i]?.layout || "텍스트형";

    // 헤더 바
    slide.addShape("rect" as PptxGenJS.ShapeType, {
      x: 0, y: 0, w: "100%", h: 0.8,
      fill: { color: C.navy }, line: { type: "none" },
    });
    slide.addText(title, {
      x: 0.5, y: 0.15, w: 12.3, h: 0.5,
      fontSize: 22, bold: true, color: C.white, fontFace: FONT,
    });

    // 본문
    const builder = LAYOUT_BUILDERS[layoutType] || buildTextSlide;
    try {
      builder(slide, sd);
    } catch {
      buildTextSlide(slide, sd);
    }

    // 푸터
    slide.addText("전략추진실", {
      x: 0.5, y: 7.1, w: 3, h: 0.3,
      fontSize: 9, color: C.darkGray, fontFace: FONT,
    });
    slide.addText(String(i + 1), {
      x: 12.0, y: 7.1, w: 0.8, h: 0.3,
      fontSize: 9, color: C.darkGray, fontFace: FONT, align: "right",
    });
  }

  // base64로 출력 (Vercel에서 파일시스템 불필요)
  const pptxOutput = await pptx.write({ outputType: "base64" });
  const pptxBase64 = typeof pptxOutput === "string" ? pptxOutput : Buffer.from(pptxOutput as ArrayBuffer).toString("base64");

  session.step = 7;

  return makeResponse(7, "done",
    `PPTX를 생성했습니다. (${slidesContent.length}장)`,
    { totalSlides: slidesContent.length },
    8, false, "PPTX 생성이 완료되었습니다.", session, pptxBase64);
}

// ─── Step 8: 최종화 ─────────────────────────────────────────────────────────

function step8(session: Session): StepResponse {
  session.step = 8;
  return makeResponse(8, "done", "장표 제작이 완료되었습니다.",
    { sessionId: session.id }, -1, false,
    "완료! 다운로드 버튼으로 PPTX를 받으세요.", session);
}

// ─── POST Handler ───────────────────────────────────────────────────────────

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { step, userInput, session: clientSession } = body as {
      step: number;
      userInput?: string;
      session?: Session;
    };

    const session: Session = clientSession || newSession();

    let result: StepResponse;

    switch (step) {
      case 1: result = await step1(session, userInput || ""); break;
      case 2: result = await step2(session, userInput || ""); break;
      case 3: result = await step3(session); break;
      case 4: result = await step4(session, userInput || ""); break;
      case 5: result = await step5(session); break;
      case 6: result = await step6(session, userInput || ""); break;
      case 7: result = await step7(session); break;
      case 8: result = step8(session); break;
      default:
        return NextResponse.json({ error: `잘못된 단계: ${step}` }, { status: 400 });
    }

    return NextResponse.json(result);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[ppt-create]", message);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
