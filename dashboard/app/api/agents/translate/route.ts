import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import PptxGenJS from "pptxgenjs";
import { resolveFile, saveTempFile } from "@/lib/file-utils";
import { readFile, writeFile } from "fs/promises";
import os from "os";
import path from "path";

const TERMINOLOGY: Record<string, string> = {
  "사업추진본부": "Business Promotion Division",
  "전략추진실": "Strategic Planning Office",
  "잡코리아": "JobKorea",
  "알바몬": "Albamon",
  "잡플래닛": "Jobplanet",
  "링커": "Linker",
  "대표이사": "CEO",
  "이사회": "Board of Directors",
  "매출액": "Revenue",
  "영업이익": "Operating Profit",
  "서치펌": "Search Firm",
  "채용대행": "RPO (Recruitment Process Outsourcing)",
  "인재검색": "Talent Search",
  "이력서": "Resume/CV",
  "공고": "Job Posting",
  "지원자": "Applicant",
  "배치": "Placement",
  "전환율": "Conversion Rate",
};

const PRESERVE_WORDS = [
  "JobKorea", "Albamon", "Jobplanet", "SEEK", "RMS", "NPS", "KPI", "OKR",
  "ARPU", "MAU", "DAU", "B2B", "B2C", "HR", "AI", "ML", "API",
  "LinkedIn", "Indeed", "Glassdoor", "CEO", "CFO", "COO", "CTO",
];

export async function POST(req: NextRequest) {
  try {
    const { pptxPath, step } = await req.json();

    if (step === 1) {
      // Step 1: Download from Blob URL → local /tmp
      const localPath = await resolveFile(pptxPath);
      const stat = await import("fs/promises").then(fs => fs.stat(localPath));
      return NextResponse.json({
        success: true, step: 1,
        output: `PPT 파일 로드 완료\n파일 크기: ${(stat.size / 1024).toFixed(1)} KB`,
        localPath, // pass to subsequent steps
      });
    }

    if (step === 2) {
      // Step 2: Extract text from PPTX using JSZip
      const localPath = await resolveFile(pptxPath);
      const JSZip = (await import("jszip")).default;
      const buffer = await readFile(localPath);
      const zip = await JSZip.loadAsync(buffer);

      const slides: { slideNum: number; texts: string[] }[] = [];
      const slideFiles = Object.keys(zip.files).filter(f => f.match(/ppt\/slides\/slide\d+\.xml$/)).sort();

      for (const slideFile of slideFiles) {
        const xml = await zip.files[slideFile].async("text");
        const texts: string[] = [];
        const regex = /<a:t>([^<]+)<\/a:t>/g;
        let match;
        while ((match = regex.exec(xml)) !== null) {
          if (match[1].trim()) texts.push(match[1]);
        }
        const slideNum = parseInt(slideFile.match(/slide(\d+)/)?.[1] || "0");
        slides.push({ slideNum, texts });
      }

      const dataPath = await saveTempFile("extracted.json", JSON.stringify(slides, null, 2));

      return NextResponse.json({
        success: true, step: 2,
        output: `텍스트 추출 완료\n슬라이드: ${slides.length}장\n텍스트 블록: ${slides.reduce((a, s) => a + s.texts.length, 0)}개`,
        dataPath, slideCount: slides.length
      });
    }

    if (step === 3) {
      return NextResponse.json({
        success: true, step: 3,
        output: `용어집 로드 완료\n매칭 용어: ${Object.keys(TERMINOLOGY).length}개\n보존어: ${PRESERVE_WORDS.length}개`
      });
    }

    if (step === 4) {
      const apiKey = process.env.ANTHROPIC_API_KEY;
      if (!apiKey) throw new Error("ANTHROPIC_API_KEY not set. Vercel 환경변수를 확인하세요.");

      const { dataPath } = await req.json().catch(() => ({ dataPath: "" }));
      // Try to find extracted data
      const tmpDir = os.tmpdir();
      const files = await import("fs/promises").then(fs => fs.readdir(tmpDir));
      const extractedFile = files.filter(f => f.includes("extracted.json")).sort().pop();
      if (!extractedFile) throw new Error("텍스트 추출 데이터가 없습니다. Step 2를 먼저 실행하세요.");

      const slides = JSON.parse(await readFile(path.join(tmpDir, extractedFile), "utf-8")) as { slideNum: number; texts: string[] }[];

      const client = new Anthropic({ apiKey });
      const translatedSlides: { slideNum: number; original: string[]; translated: string[] }[] = [];

      for (const slide of slides) {
        if (slide.texts.length === 0) {
          translatedSlides.push({ slideNum: slide.slideNum, original: [], translated: [] });
          continue;
        }

        const textsBlock = slide.texts.map((t, i) => `[${i}] ${t}`).join("\n");
        const termList = Object.entries(TERMINOLOGY).map(([k, v]) => `${k} → ${v}`).join("\n");

        const resp = await client.messages.create({
          model: "claude-sonnet-4-5-20250929",
          max_tokens: 4096,
          messages: [{
            role: "user",
            content: `Translate the following Korean presentation text to Australian English. HR/Recruitment domain.

TERMINOLOGY:
${termList}

PRESERVE: ${PRESERVE_WORDS.join(", ")}

RULES:
- Australian English (organise, colour, behaviour)
- Keep numbers/units, 억=100M/B, 조=T
- Be concise for presentation
- Already English → keep as-is

TEXT:
${textsBlock}

Return ONLY [N] translations, one per line.`
          }]
        });

        const responseText = resp.content[0].type === "text" ? resp.content[0].text : "";
        const translated: string[] = [...slide.texts];

        for (const line of responseText.split("\n")) {
          const m = line.match(/^\[(\d+)\]\s*(.+)/);
          if (m) {
            const idx = parseInt(m[1]);
            if (idx < translated.length) translated[idx] = m[2].trim();
          }
        }

        translatedSlides.push({ slideNum: slide.slideNum, original: slide.texts, translated });
      }

      const transPath = await saveTempFile("translated.json", JSON.stringify(translatedSlides, null, 2));

      return NextResponse.json({
        success: true, step: 4,
        output: `번역 완료\n${translatedSlides.length}장 슬라이드\nClaude API 호출: ${translatedSlides.filter(s => s.original.length > 0).length}회`,
        transPath
      });
    }

    if (step === 5) {
      const tmpDir = os.tmpdir();
      const files = await import("fs/promises").then(fs => fs.readdir(tmpDir));
      const transFile = files.filter(f => f.includes("translated.json")).sort().pop();
      if (!transFile) throw new Error("번역 데이터가 없습니다.");

      const transPath = path.join(tmpDir, transFile);
      const slides = JSON.parse(await readFile(transPath, "utf-8"));

      let fixes = 0;
      for (const slide of slides) {
        for (let i = 0; i < slide.translated.length; i++) {
          let t = slide.translated[i];
          t = t.replace(/\borganize\b/gi, "organise");
          t = t.replace(/\borganization\b/gi, "organisation");
          t = t.replace(/\bcolor\b/gi, "colour");
          t = t.replace(/\bbehavior\b/gi, "behaviour");
          t = t.replace(/\banalyze\b/gi, "analyse");
          t = t.replace(/\blicense\b/gi, "licence");
          if (t !== slide.translated[i]) fixes++;
          slide.translated[i] = t;
        }
      }

      await writeFile(transPath, JSON.stringify(slides, null, 2));

      return NextResponse.json({
        success: true, step: 5,
        output: `후처리 완료\n호주 영어 철자 수정: ${fixes}건`
      });
    }

    if (step === 6) {
      const tmpDir = os.tmpdir();
      const files = await import("fs/promises").then(fs => fs.readdir(tmpDir));
      const transFile = files.filter(f => f.includes("translated.json")).sort().pop();
      if (!transFile) throw new Error("번역 데이터가 없습니다.");

      const slides = JSON.parse(await readFile(path.join(tmpDir, transFile), "utf-8"));

      const pptx = new PptxGenJS();
      pptx.layout = "LAYOUT_WIDE";

      for (const slide of slides) {
        const s = pptx.addSlide();
        if (slide.translated.length === 0) continue;

        s.addText(slide.translated[0] || "", {
          x: 0.5, y: 0.3, w: "90%", h: 0.8,
          fontSize: 24, bold: true, color: "1B365D", fontFace: "Calibri",
        });

        if (slide.translated.length > 1) {
          const bodyTexts = slide.translated.slice(1).map((t: string) => ({
            text: t, options: { fontSize: 14, color: "333333", bullet: true, breakLine: true }
          }));
          s.addText(bodyTexts, {
            x: 0.5, y: 1.3, w: "90%", h: "70%", fontFace: "Calibri", valign: "top",
          });
        }
      }

      const outputPath = path.join(tmpDir, `${Date.now()}_translated.pptx`);
      await pptx.writeFile({ fileName: outputPath });

      return NextResponse.json({
        success: true, step: 6,
        output: `번역 PPT 저장 완료\n슬라이드: ${slides.length}장`,
        outputPath
      });
    }

    return NextResponse.json({ error: "Invalid step" }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
