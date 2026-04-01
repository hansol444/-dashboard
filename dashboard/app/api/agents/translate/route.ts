import { NextRequest, NextResponse } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import PptxGenJS from "pptxgenjs";

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
      // Step 1: Load PPTX and extract text
      // Since we can't easily parse PPTX in serverless, we'll use a simpler approach
      // Read the file and get basic info
      const fs = await import("fs/promises");
      const stat = await fs.stat(pptxPath);
      return NextResponse.json({
        success: true, step: 1,
        output: `PPT 파일 로드 완료\n파일 크기: ${(stat.size / 1024).toFixed(1)} KB`
      });
    }

    if (step === 2) {
      // Step 2: Extract text from PPTX using JSZip
      const fs = await import("fs/promises");
      const JSZip = (await import("jszip")).default;
      const buffer = await fs.readFile(pptxPath);
      const zip = await JSZip.loadAsync(buffer);

      const slides: { slideNum: number; texts: string[] }[] = [];
      const slideFiles = Object.keys(zip.files).filter(f => f.match(/ppt\/slides\/slide\d+\.xml$/)).sort();

      for (const slideFile of slideFiles) {
        const xml = await zip.files[slideFile].async("text");
        // Extract text between <a:t> tags
        const texts: string[] = [];
        const regex = /<a:t>([^<]+)<\/a:t>/g;
        let match;
        while ((match = regex.exec(xml)) !== null) {
          if (match[1].trim()) texts.push(match[1]);
        }
        const slideNum = parseInt(slideFile.match(/slide(\d+)/)?.[1] || "0");
        slides.push({ slideNum, texts });
      }

      // Store extracted data in /tmp for next steps
      const dataPath = pptxPath.replace(/\.[^.]+$/, "_extracted.json");
      await fs.writeFile(dataPath, JSON.stringify(slides, null, 2));

      return NextResponse.json({
        success: true, step: 2,
        output: `텍스트 추출 완료\n슬라이드: ${slides.length}장\n텍스트 블록: ${slides.reduce((a, s) => a + s.texts.length, 0)}개`,
        dataPath, slideCount: slides.length
      });
    }

    if (step === 3) {
      // Step 3: Load terminology
      return NextResponse.json({
        success: true, step: 3,
        output: `용어집 로드 완료\n매칭 용어: ${Object.keys(TERMINOLOGY).length}개\n보존어: ${PRESERVE_WORDS.length}개`
      });
    }

    if (step === 4) {
      // Step 4: Translate with Claude API
      const apiKey = process.env.ANTHROPIC_API_KEY;
      if (!apiKey) throw new Error("ANTHROPIC_API_KEY not set");

      const fs = await import("fs/promises");
      const dataPath = pptxPath.replace(/\.[^.]+$/, "_extracted.json");
      const slides = JSON.parse(await fs.readFile(dataPath, "utf-8")) as { slideNum: number; texts: string[] }[];

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
          model: "claude-sonnet-4-6-20250514",
          max_tokens: 4096,
          messages: [{
            role: "user",
            content: `Translate the following Korean presentation text to Australian English. This is HR/Recruitment domain content.

TERMINOLOGY (must use these exact translations):
${termList}

PRESERVE these words as-is: ${PRESERVE_WORDS.join(", ")}

RULES:
- Use Australian English spelling (organise, colour, behaviour, etc.)
- Keep numbers and units intact
- 억 = 100M or B (context-dependent), 조 = T
- Keep bullet structure and formatting cues
- Be concise - presentation text should be brief
- If text is already English, keep it as-is

TEXT TO TRANSLATE:
${textsBlock}

Return ONLY the translations in the same [N] format, one per line. No explanations.`
          }]
        });

        const responseText = resp.content[0].type === "text" ? resp.content[0].text : "";
        const translated: string[] = [...slide.texts]; // fallback to original

        const lines = responseText.split("\n");
        for (const line of lines) {
          const m = line.match(/^\[(\d+)\]\s*(.+)/);
          if (m) {
            const idx = parseInt(m[1]);
            if (idx < translated.length) translated[idx] = m[2].trim();
          }
        }

        translatedSlides.push({ slideNum: slide.slideNum, original: slide.texts, translated });
      }

      const transPath = pptxPath.replace(/\.[^.]+$/, "_translated.json");
      await fs.writeFile(transPath, JSON.stringify(translatedSlides, null, 2));

      return NextResponse.json({
        success: true, step: 4,
        output: `번역 완료\n${translatedSlides.length}장 슬라이드 번역\nClaude API 호출: ${translatedSlides.filter(s => s.original.length > 0).length}회`,
        transPath
      });
    }

    if (step === 5) {
      // Step 5: Post-processing
      const fs = await import("fs/promises");
      const transPath = pptxPath.replace(/\.[^.]+$/, "_translated.json");
      const slides = JSON.parse(await fs.readFile(transPath, "utf-8"));

      let fixes = 0;
      for (const slide of slides) {
        for (let i = 0; i < slide.translated.length; i++) {
          let t = slide.translated[i];
          // Australian English fixes
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

      await fs.writeFile(transPath, JSON.stringify(slides, null, 2));

      return NextResponse.json({
        success: true, step: 5,
        output: `후처리 완료\n호주 영어 철자 수정: ${fixes}건`
      });
    }

    if (step === 6) {
      // Step 6: Generate translated PPTX
      const fs = await import("fs/promises");
      const transPath = pptxPath.replace(/\.[^.]+$/, "_translated.json");
      const slides = JSON.parse(await fs.readFile(transPath, "utf-8"));

      const pptx = new PptxGenJS();
      pptx.layout = "LAYOUT_WIDE"; // 16:9

      for (const slide of slides) {
        const s = pptx.addSlide();
        if (slide.translated.length === 0) continue;

        // First text as title
        s.addText(slide.translated[0] || "", {
          x: 0.5, y: 0.3, w: "90%", h: 0.8,
          fontSize: 24, bold: true, color: "1B365D",
          fontFace: "Calibri",
        });

        // Rest as body text
        if (slide.translated.length > 1) {
          const bodyTexts = slide.translated.slice(1).map((t: string) => ({
            text: t, options: { fontSize: 14, color: "333333", bullet: true, breakLine: true }
          }));
          s.addText(bodyTexts, {
            x: 0.5, y: 1.3, w: "90%", h: "70%",
            fontFace: "Calibri", valign: "top",
          });
        }
      }

      const outputPath = `/tmp/${Date.now()}_translated.pptx`;
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
