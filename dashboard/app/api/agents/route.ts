import { NextRequest, NextResponse } from "next/server";
import { exec } from "child_process";
import { promisify } from "util";
import path from "path";

const execAsync = promisify(exec);

/**
 * 에이전트 실행 API
 * POST /api/agents
 * body: { topicFile: string, params?: Record<string, string> }
 *
 * 로컬(npm run dev)에서 실행 시 실제 스크립트 실행.
 * Vercel 배포 환경에서는 로컬 파일 접근 불가로 실행 안 됨.
 */

// 워크스페이스 루트 (dashboard 기준 상위)
const WORKSPACE = path.resolve(process.cwd(), "..");
const PLACEMENT_DIR = "C:\\Users\\ugin35\\Desktop\\Placement survey 자동화 revive";
const PPT_MAKER_DIR = path.join(WORKSPACE, "ppt-maker");

const AGENT_CONFIGS: Record<string, {
  command: string;
  cwd: string;
  description: string;
  steps: string[];
}> = {
  "regular/macro-update": {
    command: "python update_macro.py",
    cwd: WORKSPACE,
    description: "KOSIS 데이터 → Macro Analysis 엑셀 업데이트",
    steps: ["KOSIS 파일 탐색", "데이터 읽기", "Macro 엑셀 열기", "10개 시트 업데이트", "저장"],
  },
  "regular/placement-analysis": {
    command: "python run_placement_agent.py",
    cwd: PLACEMENT_DIR,
    description: "Placement Survey 전체 파이프라인 (JK+AM): R_통합 → RMS → PPT 자동 생성",
    steps: [
      "Stage 1-JK: run_jk.py → R_통합 시트 생성",
      "Stage 1-AM: run_am.py → R_통합 시트 생성",
      "Stage 2-JK: calc_rms.py → RMS 계산 Excel",
      "Stage 2-AM: calc_rms_am.py → RMS 계산 Excel",
      "Stage 3: gen_ppt.py → 분기 PPT 자동 생성",
    ],
  },
  "fluid/ppt-translate": {
    command: "python ppt-translater/translate.py",
    cwd: WORKSPACE,
    description: "한글 PPT → 영문 PPT 번역",
    steps: ["PPT 파일 로드", "텍스트박스 크기 분석", "용어집 로드", "Claude API 번역", "번역 PPT 저장"],
  },
  "fluid/budget-draft": {
    command: "node auto.js",
    cwd: WORKSPACE,
    description: "NL → 전자결재 자동 입력 (Playwright)",
    steps: ["NL 파싱 (Claude)", "문서유형 판단", "템플릿 로드", "폼 자동 입력", "검수 대기"],
  },
  "fluid/ppt-create": {
    command: "python create.py",
    cwd: PPT_MAKER_DIR,
    description: "맥락 → 스토리라인 → 레이아웃 → PPTX 생성",
    steps: ["입력 분석", "명확화 질문", "헤드메시지 초안", "레이아웃 매칭", "PPTX 생성"],
  },
  "fluid/meeting-notes": {
    command: "python meeting-notes/summarize.py",
    cwd: WORKSPACE,
    description: "TXT 녹취록 → Claude 구조화 요약 → Notion 등록",
    steps: ["TXT 파일 읽기", "Claude 요약 생성", "업무지시 추출", "Notion 등록", "GitHub 저장"],
  },
};

const IS_VERCEL = !!process.env.VERCEL;

export async function POST(request: NextRequest) {
  const body = await request.json();
  const { topicFile, params } = body as { topicFile: string; params?: Record<string, string> };

  const config = AGENT_CONFIGS[topicFile];
  if (!config) {
    return NextResponse.json(
      { error: `실행 가능한 에이전트 없음: ${topicFile}` },
      { status: 400 }
    );
  }

  // Vercel 환경: 직접 실행 불가 → 로컬 워커가 GitHub DB 폴링으로 처리
  if (IS_VERCEL) {
    return NextResponse.json({
      success: true,
      description: config.description,
      command: "(로컬 워커에서 실행 예정)",
      stdout: "Vercel 환경에서는 직접 실행이 불가합니다.\n로컬 PC에서 worker.js가 실행 중이면 자동으로 처리됩니다.",
      stderr: "",
    });
  }

  // 추가 파라미터 붙이기
  let command = config.command;
  if (params) {
    // budget-draft의 auto.js는 평문 인자를 받음 (--key 형태 아님)
    if (topicFile === "fluid/budget-draft" && params.message) {
      command = `${command} "${params.message}"`;
    } else if (topicFile === "fluid/ppt-create" && params.session) {
      // ppt-create 세션 모드: --session <id> --step <n> --input "text"
      command = `${command} --session "${params.session}" --step ${params.step || "1"}`;
      if (params.input) command += ` --input "${params.input}"`;
    } else if (topicFile === "fluid/ppt-create" && params.input && !params.session) {
      // ppt-create 단발 모드: --input "text" --output "path"
      command = `${command} --input "${params.input}" --output "output/result_${Date.now()}.pptx"`;
    } else {
      // 플래그 전용 파라미터 (값 없이 --key만 붙이는 경우)와 일반 파라미터 분리
      const flags = Object.entries(params).filter(([, v]) => v === "__flag__").map(([k]) => `--${k}`);
      const extra = Object.entries(params)
        .filter(([, v]) => v && v !== "__flag__")
        .map(([k, v]) => `--${k} "${v}"`)
        .join(" ");
      const flagStr = flags.join(" ");
      if (extra || flagStr) command = `${command} ${extra} ${flagStr}`.replace(/\s+/g, " ").trim();
    }
  }

  try {
    const { stdout, stderr } = await execAsync(command, {
      cwd: config.cwd,
      timeout: 300_000, // 5분
      env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
      encoding: "utf-8",
    });

    return NextResponse.json({
      success: true,
      description: config.description,
      command,
      stdout: stdout.slice(0, 2000),
      stderr: stderr.slice(0, 500),
    });
  } catch (err: unknown) {
    const error = err as { stdout?: string; stderr?: string; message?: string };
    return NextResponse.json({
      success: false,
      command,
      stdout: error.stdout?.slice(0, 2000) ?? "",
      stderr: error.stderr?.slice(0, 500) ?? "",
      error: error.message ?? "실행 실패",
    });
  }
}

export async function GET() {
  return NextResponse.json({
    agents: Object.entries(AGENT_CONFIGS).map(([topicFile, cfg]) => ({
      topicFile,
      command: cfg.command,
      description: cfg.description,
    })),
  });
}
