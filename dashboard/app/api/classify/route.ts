import { NextRequest, NextResponse } from "next/server";

/**
 * /api/classify
 *
 * 태스크 메시지를 Claude API로 분류해서 카테고리, 진행 단계 반환.
 *
 * 필요한 환경변수:
 *   ANTHROPIC_API_KEY — Anthropic Console > API Keys
 *
 * POST body: { message: string }
 * Response: { category, topicFile, autoLevel, steps, guide }
 */

// ─── 라우팅 테이블 (Context/Topics/index.md 기반) ───

const TOPICS: {
  keywords: string[];
  category: string;
  topicFile: string;
  autoLevel: "auto" | "manual" | "knowledge";
  guide: string;
  steps: string[];
}[] = [
  {
    keywords: ["매크로", "KOSIS", "경제지표", "선행지표", "엑셀 시트"],
    category: "Macro 업데이트",
    topicFile: "regular/macro-update.md",
    autoLevel: "auto",
    guide: "python update_macro.py",
    steps: ["KOSIS 파일 탐색", "데이터 읽기", "Macro 엑셀 열기", "10개 시트 업데이트", "저장"],
  },
  {
    keywords: ["매크로 지표", "신규 지표", "지표 발굴", "지표 추가"],
    category: "Macro 지표 발굴",
    topicFile: "regular/macro-indicators.md",
    autoLevel: "manual",
    guide: "창준님과 방향 얼라인 후 KOSIS 탐색",
    steps: ["창준님 방향 확인", "KOSIS 지표 탐색", "데이터 구조 확인", "엑셀 시트 추가"],
  },
  {
    keywords: ["플레이스먼트 서베이 업데이트", "쿼터 변경", "설문 변경", "서베이 업데이트"],
    category: "Placement Survey 업데이트",
    topicFile: "regular/placement-update.md",
    autoLevel: "manual",
    guide: "엠브레인 문주원님께 변경 요청",
    steps: ["변경 내용 정리", "엠브레인 문주원님 요청", "변경 완료 확인", "내부 기록"],
  },
  {
    keywords: ["플레이스먼트 분석", "RMS", "Cubicle", "서베이 분석", "플레이스먼트"],
    category: "Placement Survey 분석",
    topicFile: "regular/placement-analysis.md",
    autoLevel: "auto",
    guide: "run_jk.py → calc_rms.py → gen_ppt.py",
    steps: ["Raw 데이터 로드", "분류표 매칭 (run_jk.py)", "미분류 확인", "RMS 계산 (calc_rms.py)", "PPT 생성 (gen_ppt.py)"],
  },
  {
    keywords: ["장표", "PPT", "덱", "슬라이드", "번역", "영문", "translate"],
    category: "장표 제작/번역",
    topicFile: "fluid/ppt-work.md",
    autoLevel: "auto",
    guide: "장표: Claude/Genspark | 번역: python ppt-translate/translate.py",
    steps: ["스토리라인 구성", "슬라이드 제작", "데이터 삽입", "디자인 적용", "검토 후 저장"],
  },
  {
    keywords: ["회의록", "싱크", "미팅노트", "녹취록"],
    category: "회의록 정리",
    topicFile: "fluid/meeting-notes.md",
    autoLevel: "auto",
    guide: "python meeting-notes/summarize.py",
    steps: ["TXT 파일 넣기", "summarize.py 실행", "요약 검토", "Notion 등록"],
  },
  {
    keywords: ["산학협력", "EGI", "MCSA", "기프티콘", "네이버페이", "계약", "프리랜서"],
    category: "산학협력/기프티콘",
    topicFile: "fluid/academia-contract.md",
    autoLevel: "manual",
    guide: "예산품의 → 구매검토 → 구매품의 → 인장 → 계약 체결",
    steps: ["예산품의 작성", "구매검토 요청 (총무팀 이민희님)", "구매품의 작성", "인장 날인", "계약 체결"],
  },
  {
    keywords: ["예산 개념", "코스트센터", "GL계정", "품의 기초"],
    category: "예산 101",
    topicFile: "budget/budget-101.md",
    autoLevel: "knowledge",
    guide: "budget-101.md 참고",
    steps: [],
  },
  {
    keywords: ["플레이스먼트 컨커", "서베이 Concur", "엠브레인 Concur"],
    category: "Placement Survey 컨커",
    topicFile: "budget/placement-concur.md",
    autoLevel: "manual",
    guide: "예산품의 → 구매품의 → Concur",
    steps: ["예산품의 작성", "구매품의 작성", "Concur 처리", "송장 첨부", "제출"],
  },
  {
    keywords: ["엔코라인", "통역"],
    category: "엔코라인 컨커",
    topicFile: "budget/enkoline-concur.md",
    autoLevel: "manual",
    guide: "Concur → 송장 첨부 → 제출",
    steps: ["Concur Report 생성", "필드 입력", "송장 첨부", "제출"],
  },
  {
    keywords: ["컨설팅", "BCG", "자문료"],
    category: "컨설팅 자문료 컨커",
    topicFile: "budget/consulting-concur.md",
    autoLevel: "manual",
    guide: "Concur → 계약서+송장 첨부 → 제출",
    steps: ["Concur Report 생성", "필드 입력 (업무명/비고)", "계약서+송장 첨부", "제출"],
  },
  {
    keywords: ["법무법인", "LAB", "태평양 월정액"],
    category: "법무법인 비용 처리",
    topicFile: "budget/law-firm.md",
    autoLevel: "manual",
    guide: "청구서 수령 → Concur 처리",
    steps: ["청구서 수령", "Concur Report 생성", "세금계산서 첨부", "제출"],
  },
  {
    keywords: ["나인하이어", "에스크로", "스톡옵션", "매매대금"],
    category: "나인하이어 지급",
    topicFile: "budget/ninehire.md",
    autoLevel: "manual",
    guide: "에스크로 수수료 → 주식매매대금 → 스톡옵션 순서 진행",
    steps: ["에스크로 수수료 지급 (매년 12월)", "재직 여부 확인", "주식매매대금 지급", "스톡옵션 보상 지급"],
  },
  {
    keywords: ["ATS 기프티콘", "기프티콘 구매", "리워드"],
    category: "ATS 기프티콘",
    topicFile: "budget/gifticon.md",
    autoLevel: "manual",
    guide: "예산품의 → 구매품의 → KT alpha 강석현님 발송 요청",
    steps: ["예산품의 작성", "구매품의 작성", "KT alpha 발송 요청", "발송 확인", "히스토리 기록"],
  },
  {
    keywords: ["추가 예산", "예산 이월", "잔액 부족", "예산 초과"],
    category: "추가 예산 품의",
    topicFile: "budget/budget-transfer.md",
    autoLevel: "manual",
    guide: "창준님 얼라인 → 예산품의 (증액 사유 명확히)",
    steps: ["창준님 사전 얼라인", "예산품의 작성 (사유 명확히)", "첨부서류 준비", "제출"],
  },
  {
    keywords: ["신규 공급사", "벤더 등록", "공급사 등록"],
    category: "신규 공급사 등록",
    topicFile: "budget/vendor-registration.md",
    autoLevel: "knowledge",
    guide: "사업자등록증+통장 수령 → 포탈 등록 신청",
    steps: ["서류 수령 (사업자등록증, 통장)", "포탈 공급사 등록 신청", "재무회계팀 검토 확인"],
  },
];

const ROUTING_TABLE_TEXT = TOPICS.map((t) =>
  `- ${t.category} (${t.topicFile}): 키워드 예시 [${t.keywords.join(", ")}]`
).join("\n");

// ─── Claude API 호출 ───

async function classifyWithClaude(message: string): Promise<{
  category: string;
  topicFile: string;
  autoLevel: "auto" | "manual" | "knowledge";
  guide: string;
  steps: string[];
} | null> {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null;

  const prompt = `당신은 전략추진실 업무 분류 시스템입니다.
아래 업무 메시지를 읽고, 가장 적합한 카테고리 하나를 선택하세요.

## 라우팅 테이블
${ROUTING_TABLE_TEXT}

## 업무 메시지
"${message}"

## 응답 형식 (JSON만 출력, 다른 텍스트 없이)
{"topicFile": "regular/macro-update.md"}

topicFile은 반드시 위 라우팅 테이블에 있는 값 중 하나여야 합니다.
어디에도 해당하지 않으면 {"topicFile": "none"}을 반환하세요.`;

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 64,
      messages: [{ role: "user", content: prompt }],
    }),
  });

  if (!res.ok) return null;

  const data = await res.json();
  const text: string = data?.content?.[0]?.text ?? "";

  try {
    const parsed = JSON.parse(text.trim());
    const matched = TOPICS.find((t) => t.topicFile === parsed.topicFile);
    if (!matched) return null;
    return {
      category: matched.category,
      topicFile: matched.topicFile,
      autoLevel: matched.autoLevel,
      guide: matched.guide,
      steps: matched.steps,
    };
  } catch {
    return null;
  }
}

// ─── 폴백: 키워드 매칭 ───

function classifyByKeyword(message: string) {
  for (const topic of TOPICS) {
    for (const kw of topic.keywords) {
      if (message.includes(kw)) {
        return topic;
      }
    }
  }
  return null;
}

// ─── 라우트 핸들러 ───

export async function POST(request: NextRequest) {
  const { message } = await request.json();
  if (!message) return NextResponse.json({ error: "message required" }, { status: 400 });

  // Claude API 분류 시도
  const claudeResult = await classifyWithClaude(message);
  if (claudeResult) return NextResponse.json(claudeResult);

  // 폴백: 키워드 매칭
  const kwResult = classifyByKeyword(message);
  if (kwResult) {
    return NextResponse.json({
      category: kwResult.category,
      topicFile: kwResult.topicFile,
      autoLevel: kwResult.autoLevel,
      guide: kwResult.guide,
      steps: kwResult.steps,
    });
  }

  // 미분류
  return NextResponse.json({
    category: "미분류",
    topicFile: "none",
    autoLevel: "manual",
    guide: "매뉴얼에 없는 업무입니다. 창준님께 확인하세요.",
    steps: [],
  });
}
