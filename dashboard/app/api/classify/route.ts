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

/**
 * 매칭 원칙:
 * 1. 구체적인 키워드(업데이트/update 포함)가 앞에 위치 → 먼저 매칭
 * 2. 광범위한 키워드(주제명만)가 뒤에 위치 → 나머지 전부 흡수
 * 3. 대소문자 무시, 한국어·영어 모두 허용
 */
const TOPICS: {
  keywords: string[];
  category: string;
  topicFile: string;
  autoLevel: "auto" | "manual" | "knowledge";
  guide: string;
  steps: string[];
}[] = [
  // ── Macro: 업데이트 먼저, 분석은 나머지 전부 ──
  {
    keywords: [
      "매크로 업데이트", "macro update", "kosis 업데이트", "kosis update",
      "엑셀 업데이트", "excel update", "시트 업데이트",
    ],
    category: "Macro Analysis 분석",
    topicFile: "regular/macro-update.md",
    autoLevel: "auto",
    guide: "python update_macro.py → 장표 업데이트 → 창준님 공유",
    steps: [
      "업데이트 주기 확인 (월/분기/반기)",
      "데이터 소스 접속 (KOSIS 또는 고용노동통계)",
      "Raw Sheet에 데이터 입력",
      "Overview 시트 자동 반영 확인",
      "장표 복사 후 수치 + Implication 업데이트",
      "창준님께 슬랙으로 공유",
    ],
  },
  {
    keywords: [
      "매크로", "macro", "매크로 분석", "macro analysis",
      "경제지표", "economic indicator", "선행지표", "지표", "indicator",
      "kosis", "신규 지표", "지표 발굴",
    ],
    category: "Macro Analysis Update",
    topicFile: "regular/macro-indicators.md",
    autoLevel: "manual",
    guide: "메시지 구조화 → KOSIS + Claude 탐색 → 창준님 보고",
    steps: [
      "전달하고 싶은 메시지 구조화",
      "지표 후보 우선순위 정리 (Flow vs Stock, 상용/임시 구분)",
      "KOSIS + Claude로 지표 탐색",
      "대표성·메시지 전달력 평가",
      "기존 지표와 수치 정합성 확인",
      "창준님께 보고",
    ],
  },

  // ── Placement: 업데이트 먼저, 분석은 나머지 전부 ──
  {
    keywords: [
      "플레이스먼트 업데이트", "placement update",
      "서베이 업데이트", "survey update",
      "설문 변경", "쿼터 변경", "quota",
    ],
    category: "Placement Survey 업데이트",
    topicFile: "regular/placement-update.md",
    autoLevel: "manual",
    guide: "엠브레인 문주원님께 변경 요청",
    steps: [
      "직전 분기 데이터 분석 (표본 효과 vs 질문 변경 효과 분리)",
      "변경 필요 항목 정리",
      "엠브레인 문주원님께 변경 사항 전달",
      "변경 완료 확인",
      "Q별 변경 이력 내부 기록",
    ],
  },
  {
    keywords: [
      "플레이스먼트", "placement",
      "플레이스먼트 분석", "placement analysis",
      "placement survey", "서베이 분석", "survey analysis",
      "rms", "cubicle",
    ],
    category: "Placement Survey 분석",
    topicFile: "regular/placement-analysis.md",
    autoLevel: "auto",
    guide: "run_jk.py → calc_rms.py → gen_ppt.py",
    steps: [
      "엠브레인 Raw Data 수령 (JK + AM Excel)",
      "Stage 1: R_통합 생성 (run_jk.py / run_am.py)",
      "Stage 2: RMS 계산 (calc_rms.py / calc_rms_am.py)",
      "Stage 3: PPT 자동 생성 (gen_ppt.py)",
      "수작업 보완 (목차·Scatter·Appendix)",
      "창준님께 공유",
    ],
  },

  // ── 장표 번역 먼저 (더 구체적) ──
  {
    keywords: [
      "번역", "translation", "translate", "영문", "english", "영어",
      "ppt 번역", "장표 번역",
    ],
    category: "장표 번역",
    topicFile: "fluid/ppt-translate.md",
    autoLevel: "auto",
    guide: "대시보드 Agent 실행 → input/ 폴더에 PPTX 저장 → 자동 번역 → 오역 검수 → 창준님 보고",
    steps: [
      "대시보드에서 장표 번역 Agent 켜기",
      "한글 PPTX를 ppt-translate/input/ 폴더에 저장",
      "Agent가 translate.py 실행 → 후처리 7개 규칙 자동 적용 (억/조 단위·통화 순서·호주영어 등)",
      "단위·용어 오역 수동 검수 (억→0.1B KRW, 파도급→Sub-contracting, 알바천국→AH)",
      "반복 오류는 terminology.json에 추가 학습",
      "output/ 폴더의 번역본 창준님께 공유",
    ],
  },
  // ── 장표 제작 ──
  {
    keywords: [
      "장표", "ppt", "덱", "슬라이드", "deck", "slide", "presentation",
      "장표 제작", "ppt 제작",
    ],
    category: "장표 제작",
    topicFile: "fluid/ppt-create.md",
    autoLevel: "auto",
    guide: "대시보드 Agent 실행 → 입력 제공 → 헤드메시지 초안·확정 → 레이아웃 확정 → WXP PPTX 제작 → 창준님 보고",
    steps: [
      "대시보드에서 장표 제작 Agent 켜기",
      "입력 제공: 트랜스크립트 / 정리된 초안 / 대화 중 택1",
      "청중·목적·현재상태 3가지 명확화 (브리핑 질문 3개 필수)",
      "Claude가 장별 헤드메시지 초안 제시 → 9개 체크리스트 피드백 반영 → 확정",
      "장별 레이아웃 자동 매칭 (28개 유형) → 한 장씩 OK 또는 수정",
      "WXP 양식 기반 PPTX 제작 (헤더/푸터/로고/폰트/색상 자동 적용)",
      "Finalize → 창준님 보고",
    ],
  },

  // ── 회의록 ──
  {
    keywords: ["회의록", "싱크", "미팅노트", "녹취록", "meeting note", "meeting notes", "sync"],
    category: "회의록 정리",
    topicFile: "fluid/meeting-notes.md",
    autoLevel: "auto",
    guide: "input/ 폴더에 TXT 넣기 → summarize.py 실행 → 요약 검토 → Notion 등록",
    steps: [
      "input/ 폴더에 TXT 녹취록 저장",
      "python meeting-notes/summarize.py input/파일.txt 실행",
      "구조화 요약·업무지시·방향성 변화 검토",
      "Notion Meeting Notes DB 등록 (--notion 플래그)",
      "pending_actions.json → 대시보드 업무 대기목록 연동",
    ],
  },

  // ── 산학협력/계약 ──
  {
    keywords: [
      "산학협력", "academia", "egi", "mcsa",
      "기프티콘", "gifticon", "네이버페이", "naverpay",
      "계약", "contract", "프리랜서", "freelance",
    ],
    category: "산학협력/기프티콘",
    topicFile: "fluid/academia-contract.md",
    autoLevel: "manual",
    guide: "계약서 수령 → 예산품의 → 구매검토(이민희님) → 구매품의 → 계약 체결 → 송장(Concur)",
    steps: [
      "계약서 초안 수령 및 내용 확인",
      "예산품의 작성",
      "구매검토 요청 (총무팀 이민희님)",
      "구매품의 작성",
      "인장 날인 후 계약 체결",
      "송장 수령 → Concur 처리",
    ],
  },

  // ── 예산 ──
  {
    keywords: ["예산 개념", "코스트센터", "cost center", "gl계정", "품의 기초"],
    category: "예산 101",
    topicFile: "budget/budget-101.md",
    autoLevel: "knowledge",
    guide: "budget-101.md 참고",
    steps: [],
  },
  {
    keywords: ["플레이스먼트 컨커", "서베이 concur", "엠브레인 concur", "placement concur"],
    category: "Placement Survey 컨커",
    topicFile: "budget/placement-concur.md",
    autoLevel: "manual",
    guide: "예산품의 → 구매품의 → Concur",
    steps: ["예산품의 작성", "구매품의 작성", "Concur 처리", "송장 첨부", "제출"],
  },
  {
    keywords: ["엔코라인", "enkoline", "통역", "interpretation"],
    category: "엔코라인 컨커",
    topicFile: "budget/enkoline-concur.md",
    autoLevel: "manual",
    guide: "Concur → 송장 첨부 → 제출",
    steps: ["Concur Report 생성", "필드 입력", "송장 첨부", "제출"],
  },
  {
    keywords: ["컨설팅", "consulting", "bcg", "자문료"],
    category: "컨설팅 자문료 컨커",
    topicFile: "budget/consulting-concur.md",
    autoLevel: "manual",
    guide: "Concur → 계약서+송장 첨부 → 제출",
    steps: ["Concur Report 생성", "필드 입력 (업무명/비고)", "계약서+송장 첨부", "제출"],
  },
  {
    keywords: ["법무법인", "law firm", "lab", "태평양", "월정액"],
    category: "법무법인 비용 처리",
    topicFile: "budget/law-firm.md",
    autoLevel: "manual",
    guide: "청구서 수령 → Concur 처리",
    steps: ["청구서 수령", "Concur Report 생성", "세금계산서 첨부", "제출"],
  },
  // ── 나인하이어 (구체적인 것 먼저) ──
  {
    keywords: ["스톡옵션", "stock option", "스톡 옵션", "주식보상"],
    category: "나인하이어 스톡옵션",
    topicFile: "budget/ninehire-stock.md",
    autoLevel: "manual",
    guide: "스톡옵션 보상 지급",
    steps: ["지급 대상 확인", "스톡옵션 보상 계산", "지급 처리", "내부 기록"],
  },
  {
    keywords: ["주식매매대금", "매매대금", "주식 매매", "shares", "주식 지급"],
    category: "나인하이어 주식매매대금",
    topicFile: "budget/ninehire-shares.md",
    autoLevel: "manual",
    guide: "재직 여부 확인 후 주식매매대금 지급",
    steps: ["재직 여부 확인", "지급 금액 확인", "주식매매대금 지급", "내부 기록"],
  },
  {
    keywords: ["나인하이어", "ninehire", "에스크로", "escrow", "수수료"],
    category: "나인하이어 에스크로 수수료",
    topicFile: "budget/ninehire-escrow.md",
    autoLevel: "manual",
    guide: "에스크로 수수료 지급 (매년 12월)",
    steps: ["수수료 청구서 수령", "금액 확인", "지급 처리", "내부 기록"],
  },
  {
    keywords: ["ats 기프티콘", "기프티콘 구매", "리워드", "reward", "kt alpha"],
    category: "ATS 기프티콘",
    topicFile: "budget/gifticon.md",
    autoLevel: "manual",
    guide: "예산품의 → 구매품의 → KT alpha 강석현님 발송 요청",
    steps: ["예산품의 작성", "구매품의 작성", "KT alpha 발송 요청", "발송 확인", "히스토리 기록"],
  },
  {
    keywords: ["추가 예산", "예산 이월", "budget transfer", "잔액 부족", "예산 초과", "budget overrun"],
    category: "추가 예산 품의",
    topicFile: "budget/budget-transfer.md",
    autoLevel: "manual",
    guide: "창준님 얼라인 → 예산품의 (증액 사유 명확히)",
    steps: ["창준님 사전 얼라인", "예산품의 작성 (사유 명확히)", "첨부서류 준비", "제출"],
  },
  {
    keywords: ["신규 공급사", "벤더 등록", "공급사 등록", "vendor", "vendor registration", "supplier"],
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
  const lower = message.toLowerCase();
  for (const topic of TOPICS) {
    for (const kw of topic.keywords) {
      if (lower.includes(kw.toLowerCase())) {
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
