import { NextRequest, NextResponse } from "next/server";

/**
 * 에이전트 실행 API
 * POST /api/agents
 * body: { category: string, params?: Record<string, string> }
 *
 * Vercel 환경에서는 로컬 스크립트 실행 불가.
 * 실행 가이드와 시뮬레이션 응답을 반환합니다.
 */

const AGENT_CONFIGS: Record<string, { command: string; description: string; steps: string[] }> = {
  "Macro 분석": {
    command: "python update_macro.py",
    description: "KOSIS 데이터 → Macro Analysis 엑셀 업데이트",
    steps: ["KOSIS 파일 탐색", "데이터 읽기", "Macro 엑셀 열기", "10개 시트 업데이트", "저장"],
  },
  "장표 번역": {
    command: "python ppt-translate/translate.py --input input/파일.pptx",
    description: "한글 PPT → 영문 PPT 번역",
    steps: ["PPT 파일 로드", "텍스트박스 크기 분석", "용어집 로드", "Claude API 번역", "번역 PPT 저장"],
  },
  "예산·구매 품의": {
    command: "node fill.js survey-b --dry-run",
    description: "전자결재 양식 자동 입력",
    steps: ["문서유형 판단", "템플릿 로드", "폼 자동 입력 (dry-run)", "사용자 검수", "실제 제출"],
  },
  "회의록 생성": {
    command: "python meeting-notes/summarize.py input/회의록.txt --notion",
    description: "TXT → Claude 요약 → Notion 등록",
    steps: ["TXT 파일 읽기", "구조화 요약 생성", "업무지시 추출", "프리뷰 생성", "Notion 등록"],
  },
};

export async function POST(request: NextRequest) {
  const body = await request.json();
  const { category } = body;

  const config = AGENT_CONFIGS[category];
  if (!config) {
    return NextResponse.json(
      { error: `지원하지 않는 카테고리: ${category}` },
      { status: 400 }
    );
  }

  // 로컬 실행 환경에서만 실제 스크립트 실행 가능
  // Vercel에서는 가이드 + 시뮬레이션 반환
  return NextResponse.json({
    success: true,
    description: config.description,
    command: config.command,
    steps: config.steps,
    message: "로컬 환경에서 위 명령어를 실행하세요.",
  });
}

export async function GET() {
  return NextResponse.json({
    agents: [
      { category: "Macro 분석", command: "python update_macro.py", ready: true },
      { category: "장표 번역", command: "python ppt-translate/translate.py", ready: true },
      { category: "예산·구매 품의", command: "node fill.js", ready: true },
      { category: "회의록 생성", command: "python meeting-notes/summarize.py", ready: true },
      { category: "장표 제작", command: "(미구현)", ready: false },
      { category: "Placement 분석", command: "(미구현)", ready: false },
    ],
  });
}
