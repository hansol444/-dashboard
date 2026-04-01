import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import PptxGenJS from "pptxgenjs";
import { readFile, writeFile } from "fs/promises";
import path from "path";

const SYSTEM_PROMPT = `너는 사업추진본부의 "장표 제작 도우미"야. 프레젠테이션(장표) 제작을 도와주는 역할이야.
핵심 원칙: "메시지 먼저, 비주얼 나중"
응답은 반드시 JSON 형식으로만 해.`;

interface Session {
  id: string;
  step: number;
  input?: string;
  inputType?: string;
  clarifications?: string[];
  headMessages?: { slide: number; title: string; message: string; layoutType: string }[];
  layouts?: { slide: number; type: string; confirmed: boolean }[];
  history: { role: string; content: string }[];
}

async function getSession(sessionId: string): Promise<Session> {
  try {
    const data = await readFile(path.join("/tmp", `session_${sessionId}.json`), "utf-8");
    return JSON.parse(data);
  } catch {
    return { id: sessionId, step: 0, history: [] };
  }
}

async function saveSession(session: Session) {
  await writeFile(path.join("/tmp", `session_${session.id}.json`), JSON.stringify(session, null, 2));
}

export async function POST(req: NextRequest) {
  try {
    const { sessionId, step, userInput } = await req.json();
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) throw new Error("ANTHROPIC_API_KEY not set");

    const client = new Anthropic({ apiKey });
    const session = await getSession(sessionId || `s_${Date.now()}`);

    // Step 1: Input Analysis
    if (step === 1) {
      session.input = userInput;
      session.step = 1;

      const resp = await client.messages.create({
        model: "claude-sonnet-4-6-20250514",
        max_tokens: 2048,
        system: SYSTEM_PROMPT,
        messages: [{
          role: "user",
          content: `다음 입력을 분석해. 입력 유형(transcript/draft/conversation)과 핵심 주제를 파악해.

입력:
${userInput}

JSON으로 응답:
{
  "inputType": "transcript|draft|conversation",
  "topics": ["주제1", "주제2"],
  "summary": "한줄 요약",
  "suggestedSlideCount": 숫자
}`
        }]
      });

      const text = resp.content[0].type === "text" ? resp.content[0].text : "";
      let parsed;
      try {
        const jsonMatch = text.match(/\{[\s\S]*\}/);
        parsed = jsonMatch ? JSON.parse(jsonMatch[0]) : { inputType: "draft", topics: [], summary: text, suggestedSlideCount: 5 };
      } catch {
        parsed = { inputType: "draft", topics: [], summary: text, suggestedSlideCount: 5 };
      }

      session.inputType = parsed.inputType;
      session.history.push({ role: "user", content: userInput });
      session.history.push({ role: "assistant", content: JSON.stringify(parsed) });
      await saveSession(session);

      return NextResponse.json({ success: true, step: 1, sessionId: session.id, data: parsed });
    }

    // Step 2: Clarification
    if (step === 2) {
      session.step = 2;

      const resp = await client.messages.create({
        model: "claude-sonnet-4-6-20250514",
        max_tokens: 1024,
        system: SYSTEM_PROMPT,
        messages: [{
          role: "user",
          content: `이 내용으로 장표를 만들려고 해:
${session.input}

명확화를 위한 질문 3개를 만들어. JSON으로:
{
  "questions": [
    {"id": 1, "question": "누구 입장에서 발표하나요? (경영진/실무자/외부)"},
    {"id": 2, "question": "..."},
    {"id": 3, "question": "..."}
  ]
}`
        }]
      });

      const text = resp.content[0].type === "text" ? resp.content[0].text : "";
      let parsed;
      try {
        const jsonMatch = text.match(/\{[\s\S]*\}/);
        parsed = jsonMatch ? JSON.parse(jsonMatch[0]) : { questions: [] };
      } catch {
        parsed = { questions: [{ id: 1, question: "대상 청중은 누구인가요?" }] };
      }

      session.clarifications = parsed.questions;
      await saveSession(session);

      return NextResponse.json({ success: true, step: 2, sessionId: session.id, data: parsed });
    }

    // Step 3: Head Messages Draft
    if (step === 3) {
      session.step = 3;
      // userInput contains answers to clarification questions

      const resp = await client.messages.create({
        model: "claude-sonnet-4-6-20250514",
        max_tokens: 4096,
        system: SYSTEM_PROMPT,
        messages: [{
          role: "user",
          content: `장표 내용: ${session.input}

추가 맥락: ${userInput || "없음"}

각 슬라이드의 헤드메시지를 작성해. JSON으로:
{
  "slides": [
    {"slide": 1, "title": "슬라이드 제목", "headMessage": "핵심 메시지 한 줄", "layoutType": "텍스트형|표형|차트형|다이어그램형|타임라인형|하이브리드형", "bulletPoints": ["포인트1", "포인트2"]}
  ]
}`
        }]
      });

      const text = resp.content[0].type === "text" ? resp.content[0].text : "";
      let parsed;
      try {
        const jsonMatch = text.match(/\{[\s\S]*\}/);
        parsed = jsonMatch ? JSON.parse(jsonMatch[0]) : { slides: [] };
      } catch {
        parsed = { slides: [] };
      }

      session.headMessages = parsed.slides;
      await saveSession(session);

      return NextResponse.json({ success: true, step: 3, sessionId: session.id, data: parsed });
    }

    // Step 4: User confirms head messages (just save confirmation)
    if (step === 4) {
      session.step = 4;
      if (userInput === "확정" || userInput === "OK") {
        await saveSession(session);
        return NextResponse.json({ success: true, step: 4, sessionId: session.id, data: { confirmed: true } });
      }
      // User wants modifications - re-run step 3 with feedback
      return NextResponse.json({ success: true, step: 4, sessionId: session.id, data: { confirmed: false, message: "수정사항을 반영하여 3단계를 다시 실행하세요." } });
    }

    // Step 5: Layout selection (auto from step 3 data)
    if (step === 5) {
      session.step = 5;
      const layouts = (session.headMessages || []).map(s => ({
        slide: s.slide,
        type: s.layoutType || "텍스트형",
        confirmed: false
      }));
      session.layouts = layouts;
      await saveSession(session);

      return NextResponse.json({ success: true, step: 5, sessionId: session.id, data: { layouts } });
    }

    // Step 6: Layout confirmation
    if (step === 6) {
      session.step = 6;
      if (session.layouts) {
        session.layouts = session.layouts.map(l => ({ ...l, confirmed: true }));
      }
      await saveSession(session);
      return NextResponse.json({ success: true, step: 6, sessionId: session.id, data: { confirmed: true } });
    }

    // Step 7: Generate PPTX
    if (step === 7) {
      session.step = 7;

      const pptx = new PptxGenJS();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "전략추진실 장표 제작 도우미";

      const COLORS = {
        navy: "1B365D",
        teal: "00B388",
        dark: "2D3748",
        gray: "718096",
        lightGray: "F7FAFC",
        white: "FFFFFF",
      };

      for (const slideData of (session.headMessages || [])) {
        const slide = pptx.addSlide();

        // Background
        slide.background = { color: COLORS.white };

        // Header bar
        slide.addShape("rect" as any, { x: 0, y: 0, w: "100%", h: 0.06, fill: { color: COLORS.teal } });

        // Title
        slide.addText(slideData.title || "", {
          x: 0.5, y: 0.3, w: "90%", h: 0.6,
          fontSize: 28, bold: true, color: COLORS.navy,
          fontFace: "Malgun Gothic",
        });

        // Head message
        slide.addText((slideData as any).headMessage || slideData.message || "", {
          x: 0.5, y: 0.9, w: "90%", h: 0.5,
          fontSize: 16, color: COLORS.teal, italic: true,
          fontFace: "Malgun Gothic",
        });

        // Body content based on layout type
        const bullets = (slideData as any).bulletPoints || [];
        if (bullets.length > 0) {
          const bodyTexts = bullets.map((b: string) => ({
            text: b,
            options: { fontSize: 14, color: COLORS.dark, bullet: { code: "2022" }, breakLine: true, paraSpaceAfter: 8 }
          }));
          slide.addText(bodyTexts, {
            x: 0.5, y: 1.6, w: "90%", h: "60%",
            fontFace: "Malgun Gothic", valign: "top",
          });
        }

        // Footer
        slide.addText(`슬라이드 ${slideData.slide}`, {
          x: 0.5, y: "92%", w: 2, h: 0.3,
          fontSize: 9, color: COLORS.gray, fontFace: "Malgun Gothic",
        });
      }

      const outputPath = `/tmp/${Date.now()}_presentation.pptx`;
      await pptx.writeFile({ fileName: outputPath });

      await saveSession(session);

      return NextResponse.json({
        success: true, step: 7, sessionId: session.id,
        data: { outputPath, slideCount: (session.headMessages || []).length },
        outputPath
      });
    }

    // Step 8: Finalize
    if (step === 8) {
      session.step = 8;
      await saveSession(session);
      return NextResponse.json({
        success: true, step: 8, sessionId: session.id,
        data: { message: "장표 제작이 완료되었습니다. 다운로드 버튼을 눌러주세요." }
      });
    }

    return NextResponse.json({ error: "Invalid step" }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
