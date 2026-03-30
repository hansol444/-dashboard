"use client";

import { useState, useRef, useEffect, useCallback } from "react";

// ─── 타입 ───
type TaskStatus = "pending" | "in_progress" | "running" | "done";
type AutoLevel = "auto" | "manual" | "knowledge";
type TabType = "tasks" | "meeting" | "knowledge";

interface ExecutionStep {
  label: string;
  status: "pending" | "running" | "done" | "error";
}

interface Task {
  id: string;
  from: string;
  to: string;
  message: string;
  category: string;
  deadline: string;
  status: TaskStatus;
  autoLevel: AutoLevel;
  guide: string;
  channel: string;
  timestamp: string;
  executionSteps?: ExecutionStep[];
  outputFile?: string;
}

interface MeetingSummary {
  id: string;
  fileName: string;
  date: string;
  summary: string;
  participants: string[];
  taskAssignments: { assignee: string; task: string; deadline: string }[];
  directionChanges: { fromWho: string; content: string }[];
  status: "processing" | "done" | "error";
  notionUrl?: string;
}

// ─── 라우팅 테이블 ───
const ROUTING: Record<string, { category: string; autoLevel: AutoLevel; guide: string; steps: string[] }> = {
  "플레이스먼트|RMS|서베이|26Q": {
    category: "Placement 분석", autoLevel: "auto",
    guide: "Raw → run_jk/am → calc_rms → gen_ppt",
    steps: ["Raw 데이터 로드", "분류표 매칭 (run_jk.py)", "미분류 리포트 생성", "RMS 계산 (calc_rms.py)", "PPT 생성 (gen_ppt.py)"],
  },
  "매크로|KOSIS|경제지표": {
    category: "Macro 분석", autoLevel: "auto",
    guide: "python update_macro.py",
    steps: ["KOSIS 파일 탐색", "데이터 읽기", "Macro 엑셀 열기", "10개 시트 업데이트", "저장"],
  },
  "장표|PPT|덱|슬라이드": {
    category: "장표 제작", autoLevel: "auto",
    guide: "Claude/Genspark 기반 PPT 생성",
    steps: ["스토리라인 생성", "슬라이드 구성", "데이터 삽입", "디자인 적용", "PPT 저장"],
  },
  "번역|영문|English|translate": {
    category: "장표 번역", autoLevel: "auto",
    guide: "python ppt-translate/translate.py",
    steps: ["PPT 파일 로드", "텍스트박스 크기 분석", "용어집 로드", "Claude API 번역", "후처리 적용", "번역 PPT 저장"],
  },
  "품의|예산|구매|전자결재": {
    category: "예산·구매 품의", autoLevel: "auto",
    guide: "node fill.js <문서유형> --dry-run",
    steps: ["문서유형 판단", "템플릿 로드", "폼 자동 입력 (dry-run)", "사용자 검수", "실제 제출"],
  },
  "회의록|싱크|미팅노트": {
    category: "회의록 생성", autoLevel: "auto",
    guide: "python meeting-notes/summarize.py",
    steps: ["TXT 파일 읽기", "구조화 요약 생성", "업무지시/방향성 추출", "프리뷰 생성", "Notion 등록"],
  },
  "Concur|송장|세금계산서": {
    category: "Concur 처리", autoLevel: "manual",
    guide: "담당자 변경 → 연동 송장 마무리 → 제출",
    steps: [],
  },
  "인보이스|구독|Invoice": {
    category: "인보이스 업데이트", autoLevel: "manual",
    guide: "매월 21일 각 서비스 결제 내역 확인 → 노션 정리",
    steps: [],
  },
  "기프티콘|네이버페이|리워드": {
    category: "기프티콘 발송", autoLevel: "manual",
    guide: "KT알파 견적 → 발주 → 발송 (아웃룩)",
    steps: [],
  },
  "BCG|컨설팅": {
    category: "BCG 서포트", autoLevel: "manual",
    guide: "창준님 지시 기반 판단. 셰어포인트 > 8. 참고자료",
    steps: [],
  },
  "나인하이어|에스크로|스톡옵션|매매대금": {
    category: "나인하이어 지급", autoLevel: "manual",
    guide: "에스크로 수수료 → 주식매매대금 → 스톡옵션 순서 진행",
    steps: [],
  },
  "이월|잔액 부족": {
    category: "예산 이월", autoLevel: "manual",
    guide: "전자결재 > 예산배정(이월) 문건 등록",
    steps: [],
  },
  "공급사|벤더 등록": {
    category: "신규공급사 등록", autoLevel: "manual",
    guide: "사업자등록증+담당자정보 → 전자결재 등록",
    steps: [],
  },
  "산학|EGI|MCSA|프리랜서": {
    category: "산학협력 계약", autoLevel: "manual",
    guide: "예산품의→구매검토→구매품의→인장→계약 체결 (5단계)",
    steps: [],
  },
  "엔코라인|통역": {
    category: "엔코라인 통역/재계약", autoLevel: "manual",
    guide: "Concur 송장 처리 또는 재계약 (만기 1달 전 착수)",
    steps: [],
  },
  "ENS|가이드포인트": {
    category: "ENS 대금/재계약", autoLevel: "manual",
    guide: "외화 송장 3개 작성 / 재계약 프로세스",
    steps: [],
  },
  "주차|Guest": {
    category: "게스트 주차", autoLevel: "manual",
    guide: "http://112.169.105.131/login → 차량번호 등록 (1시간 무료)",
    steps: [],
  },
  "뉴스|클리핑|NotebookLM": {
    category: "뉴스 클리핑", autoLevel: "manual",
    guide: "GPT News Bot 관리 + NotebookLM daily 업로드",
    steps: [],
  },
};

// ─── 지식 베이스 (수동 업무 가이드) ───
const KNOWLEDGE_BASE = [
  {
    category: "예산 이월",
    trigger: "재무회계팀에서 '잔액 부족' 슬랙 올 때",
    steps: [
      "전자결재 > 문건등록 > [예산관리]1.예산관리 > 7)예산배정(이월) 선택",
      "시행일: 기안일과 동일",
      "첨부: 해당 건의 예산품의 문서",
    ],
    contacts: ["재무회계팀 박은미님", "경영기획팀 임장식님"],
  },
  {
    category: "신규공급사 등록",
    trigger: "새 업체와 거래 시작 시",
    steps: [
      "필요서류: 사업자등록증, 담당자정보(이름/부서/전화/이메일), 계좌사본",
      "잡코리아 포탈 > 전자결재 > 문건등록",
      "[공급사 등록/변경]1.신규공급사 등록 선택",
      "업체 등록 요청서 작성 + 서류 압축 첨부",
    ],
    contacts: ["총무팀 남영현님"],
  },
  {
    category: "산학협력 계약",
    trigger: "대학교 학회와 프로젝트 진행 시 (예: EGI, MCSA)",
    steps: [
      "1. 예산품의 작성",
      "2. 구매검토 요청 (이메일 → 총무팀 이민희님, 참조: 박현수/남영현)",
      "3. 구매품의 작성 (예산품의 하위문건)",
      "4. 인장관리 (CSO 서명 필요)",
      "5. 계약 체결",
      "※ 500만원 이상 시 입찰생략요청서 필요",
    ],
    contacts: ["총무팀 이민희님", "재무회계팀 김세진님"],
  },
  {
    category: "엔코라인 통역/재계약",
    trigger: "통역 비용 처리 또는 연간 재계약 시",
    steps: [
      "통역 Concur: 재무회계팀에 송장 담당자 변경 요청 → 연동 송장으로 처리",
      "재계약: 만기 1달 전 착수 필수!",
      "재계약 순서: 예산품의 → 구매검토(총무팀) → 구매품의",
      "견적서는 엔코라인 장시몬 본부장에게 요청 (ymlee@enkoline.com)",
    ],
    contacts: ["엔코라인 장시몬 본부장", "재무회계팀 김세진님", "총무팀 남영현님"],
  },
  {
    category: "나인하이어 지급",
    trigger: "스케줄: 2026.03.31, 2026.08.16",
    steps: [
      "에스크로 수수료 지급 먼저 (매년 12월 1천만원)",
      "그 다음 주식 매매대금 지급 (재직 여부 확인 필수: 정승현, 이예린, 최돈민, 이경환)",
      "스톡옵션 보상 지급: 김재인, 안정태, 이정욱",
      "※ 웍스피어 사명 변경으로 계약서 변경 진행 중",
    ],
    contacts: ["재무회계팀"],
  },
  {
    category: "ENS 대금/재계약",
    trigger: "가이드포인트 결제 또는 연간 재계약 시",
    steps: [
      "외화 지급의 경우 총 3개 송장 작성 필요",
      "재계약 프로세스는 엔코라인과 동일 (1달 전 착수)",
      "벤더 등록 선행 필요",
    ],
    contacts: ["재무회계팀 최홍근님"],
  },
  {
    category: "뉴스 클리핑",
    trigger: "GPT News Bot 관리 시",
    steps: [
      "Slack API (api.slack.com/apps) → GPT Newsletter Bot",
      "NotebookLM에 daily 기사 업로드",
      "링커모여라 + 전략추진실 채널 기사 정리",
    ],
    contacts: [],
  },
  {
    category: "게스트 주차",
    trigger: "외부 손님 방문 시",
    steps: [
      "http://112.169.105.131/login 접속",
      "ID: f1801 / PW: f1801",
      "차량번호 등록 → 할인 클릭",
      "※ 1시간만 무료, 출차까지 1시간 내 수행 필요",
    ],
    contacts: ["총무팀 김용헌님"],
  },
];

function matchCategory(text: string) {
  for (const [keywords, info] of Object.entries(ROUTING)) {
    const regex = new RegExp(keywords.split("|").join("|"), "i");
    if (regex.test(text)) return info;
  }
  return { category: "미분류", autoLevel: "manual" as AutoLevel, guide: "매뉴얼에 없는 업무입니다. 창준님께 확인하세요.", steps: [] };
}

// ─── 더미 데이터 ───
const INITIAL_TASKS: Task[] = [
  {
    id: "1", from: "이창준", to: "주호연",
    message: "26Q1 Placement Survey 해줘",
    category: "Placement 분석", deadline: "2026-03-28",
    status: "pending", autoLevel: "auto",
    guide: "Raw → run_jk/am → calc_rms → gen_ppt",
    channel: "#section-전략추진실-창준님", timestamp: "2026-03-25 11:10",
  },
  {
    id: "2", from: "이창준", to: "임성욱",
    message: "매크로 3월분 업데이트해줘",
    category: "Macro 분석", deadline: "2026-03-31",
    status: "in_progress", autoLevel: "auto",
    guide: "python update_macro.py",
    channel: "#section-전략추진실-창준님", timestamp: "2026-03-23 14:00",
  },
  {
    id: "3", from: "이창준", to: "주호연",
    message: "BCG 측 요청 장표 정리해서 공유해줘",
    category: "BCG 서포트", deadline: "2026-04-02",
    status: "pending", autoLevel: "manual",
    guide: "창준님 지시 기반 판단. 셰어포인트 > 8. 참고자료",
    channel: "#section-전략추진실-창준님", timestamp: "2026-03-24 16:30",
  },
];

// ─── 컴포넌트 ───
function StatusBadge({ status }: { status: TaskStatus }) {
  const styles = {
    pending: "bg-yellow-500/20 text-yellow-400 border-yellow-500/30",
    in_progress: "bg-blue-500/20 text-blue-400 border-blue-500/30",
    running: "bg-purple-500/20 text-purple-400 border-purple-500/30",
    done: "bg-green-500/20 text-green-400 border-green-500/30",
  };
  const labels = { pending: "대기", in_progress: "진행 중", running: "실행 중", done: "완료" };
  return (
    <span className={`px-2 py-0.5 text-xs rounded border ${styles[status]}`}>
      {labels[status]}
    </span>
  );
}

function AutoBadge({ level }: { level: AutoLevel }) {
  const styles = {
    auto: "bg-green-500/20 text-green-400 border-green-500/30",
    manual: "bg-yellow-500/20 text-yellow-400 border-yellow-500/30",
    knowledge: "bg-gray-500/20 text-gray-400 border-gray-500/30",
  };
  const labels = { auto: "🟢 자동", manual: "🟡 수동", knowledge: "📚 지식" };
  return (
    <span className={`px-2 py-0.5 text-xs rounded border ${styles[level]}`}>
      {labels[level]}
    </span>
  );
}

function ExecutionLog({ steps, outputFile }: { steps: ExecutionStep[]; outputFile?: string }) {
  return (
    <div className="mt-3 p-3 rounded-lg" style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
      <div className="text-xs font-medium mb-2" style={{ color: "var(--text-muted)" }}>실행 로그</div>
      <div className="space-y-1">
        {steps.map((step, i) => (
          <div key={i} className="flex items-center gap-2 text-xs">
            <span>
              {step.status === "done" && "✅"}
              {step.status === "running" && "⏳"}
              {step.status === "pending" && "⚪"}
              {step.status === "error" && "❌"}
            </span>
            <span className={step.status === "running" ? "text-purple-400" : step.status === "done" ? "text-green-400" : "text-gray-500"}>
              {step.label}
            </span>
          </div>
        ))}
      </div>
      {outputFile && (
        <div className="mt-3 pt-2" style={{ borderTop: "1px solid var(--border)" }}>
          <button className="inline-flex items-center gap-1 px-3 py-1.5 rounded text-xs font-medium text-black"
            style={{ background: "var(--accent)" }}>
            📥 결과물 다운로드 — {outputFile.split("/").pop()}
          </button>
        </div>
      )}
    </div>
  );
}

function SkillCard({ name, ready, command, onRun }: { name: string; ready: boolean; command: string; onRun?: () => void }) {
  const [running, setRunning] = useState(false);
  const [result, setResult] = useState<{ success: boolean; message: string } | null>(null);

  const handleRun = async () => {
    if (!ready || running) return;
    setRunning(true);
    setResult(null);

    try {
      const res = await fetch("/api/agents", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category: name }),
      });
      const data = await res.json();
      setResult({
        success: data.success,
        message: data.success
          ? `완료! ${data.outputPath ? `→ ${data.outputPath}` : ""}`
          : `오류: ${(data.stderr || data.error || "").slice(0, 150)}`,
      });
      if (onRun) onRun();
    } catch (err) {
      setResult({ success: false, message: `서버 연결 실패: ${err}` });
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className={`p-4 rounded-xl border transition-all ${ready ? "border-green-500/30 bg-green-500/5" : "border-red-500/30 bg-red-500/5"}`}>
      <div className="flex items-center justify-between mb-2">
        <span className="font-medium text-sm">{name}</span>
        <span className={`text-xs ${ready ? "text-green-400" : "text-red-400"}`}>
          {ready ? "✅ 구현됨" : "❌ 미구현"}
        </span>
      </div>
      <code className="text-xs text-gray-400 block truncate mb-3">{command}</code>
      {ready && (
        <button
          onClick={handleRun}
          disabled={running}
          className={`w-full px-3 py-1.5 rounded text-xs font-medium transition-all ${running ? "bg-gray-700 text-gray-400 cursor-wait" : "text-black hover:opacity-80"}`}
          style={running ? {} : { background: "var(--accent)" }}>
          {running ? "⏳ 실행 중..." : "▶ 실행"}
        </button>
      )}
      {result && (
        <div className={`mt-2 p-2 rounded text-xs ${result.success ? "bg-green-500/10 text-green-400" : "bg-red-500/10 text-red-400"}`}>
          {result.message}
        </div>
      )}
    </div>
  );
}

function TaskGuide({ task }: { task: Task }) {
  const [expanded, setExpanded] = useState(false);
  const kb = KNOWLEDGE_BASE.find((k) => k.category === task.category);

  if (!kb) {
    return (
      <div className="mt-3 p-3 rounded-lg text-xs" style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
        <span style={{ color: "var(--text-muted)" }}>진행방법: </span>
        <span>{task.guide}</span>
      </div>
    );
  }

  return (
    <div className="mt-3 rounded-lg text-xs" style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
      <div className="p-3 flex items-center justify-between cursor-pointer hover:bg-white/5 rounded-lg transition-all"
        onClick={() => setExpanded(!expanded)}>
        <div>
          <span style={{ color: "var(--text-muted)" }}>진행방법: </span>
          <span>{task.guide}</span>
        </div>
        <button className="px-2 py-0.5 rounded text-xs border border-yellow-500/30 text-yellow-400 hover:bg-yellow-500/10 transition-all shrink-0 ml-2">
          {expanded ? "접기 ▲" : "📋 가이드 보기"}
        </button>
      </div>
      {expanded && (
        <div className="px-3 pb-3 space-y-2">
          <div className="border-t border-gray-800 pt-2" />
          <div className="space-y-1">
            {kb.steps.map((step, i) => (
              <div key={i} className="flex gap-2 text-xs text-gray-300">
                <span className="text-gray-500 shrink-0">{i + 1}.</span>
                <span>{step}</span>
              </div>
            ))}
          </div>
          {kb.contacts.length > 0 && (
            <div className="pt-2 border-t border-gray-800 text-xs" style={{ color: "var(--text-muted)" }}>
              📞 담당자: {kb.contacts.join(" · ")}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function KnowledgeCard({ item }: { item: typeof KNOWLEDGE_BASE[0] }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <div className="p-4 rounded-xl border border-gray-700/50 bg-gray-900/30 hover:border-gray-600/50 transition-all">
      <div className="flex items-center justify-between cursor-pointer" onClick={() => setExpanded(!expanded)}>
        <div>
          <span className="font-medium text-sm">{item.category}</span>
          <span className="ml-2 text-xs px-2 py-0.5 rounded bg-gray-800 text-gray-400">📚 지식</span>
        </div>
        <span className="text-gray-500 text-sm">{expanded ? "▲" : "▼"}</span>
      </div>
      <div className="text-xs mt-1" style={{ color: "var(--text-muted)" }}>트리거: {item.trigger}</div>
      {expanded && (
        <div className="mt-3 space-y-3">
          <div className="p-3 rounded-lg" style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
            <div className="text-xs font-medium mb-2" style={{ color: "var(--accent)" }}>진행 절차</div>
            <div className="space-y-1">
              {item.steps.map((step, i) => (
                <div key={i} className="text-xs text-gray-300 flex gap-2">
                  <span className="text-gray-500 shrink-0">{i + 1}.</span>
                  <span>{step}</span>
                </div>
              ))}
            </div>
          </div>
          {item.contacts.length > 0 && (
            <div className="text-xs" style={{ color: "var(--text-muted)" }}>
              담당자: {item.contacts.join(" · ")}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── 메인 대시보드 ───
export default function Dashboard() {
  const [tasks, setTasks] = useState<Task[]>(INITIAL_TASKS);
  const [newMessage, setNewMessage] = useState("");
  const [newFrom, setNewFrom] = useState("");
  const [newTo, setNewTo] = useState("주호연");
  const [newDeadline, setNewDeadline] = useState("");
  const [filter, setFilter] = useState<"all" | "pending" | "in_progress" | "running" | "done">("all");
  const [activeTab, setActiveTab] = useState<TabType>("tasks");
  const [meetings, setMeetings] = useState<MeetingSummary[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // GitHub DB에서 전체 태스크 폴링 (30초마다)
  const pollTasks = useCallback(async () => {
    try {
      const res = await fetch("/api/tasks");
      const data = await res.json();
      if (!data.tasks) return;
      const apiTasks: Task[] = data.tasks.map((t: { id: string; from: string; to: string; message: string; channel: string; timestamp: string; deadline?: string; status?: TaskStatus }) => {
        const matched = matchCategory(t.message);
        return {
          id: t.id,
          from: t.from,
          to: t.to,
          message: t.message,
          category: matched.category,
          deadline: t.deadline || "미정",
          status: t.status || "pending" as TaskStatus,
          autoLevel: matched.autoLevel,
          guide: matched.guide,
          channel: t.channel,
          timestamp: t.timestamp,
        };
      });
      // GitHub 태스크로 전체 교체 (수동 입력 태스크는 유지)
      setTasks((prev) => {
        const manual = prev.filter((t) => t.channel === "수동 입력");
        const apiIds = new Set(apiTasks.map((t) => t.id));
        const freshManual = manual.filter((t) => !apiIds.has(t.id));
        return [...apiTasks, ...freshManual];
      });
    } catch {
      // 서버 미연결 시 무시
    }
  }, []);

  useEffect(() => {
    pollTasks();
    const interval = setInterval(pollTasks, 30000); // 30초마다 폴링
    return () => clearInterval(interval);
  }, [pollTasks]);

  const addTask = async () => {
    if (!newMessage.trim()) return;
    const matched = matchCategory(newMessage);
    const task: Task = {
      id: `manual-${Date.now()}`,
      from: newFrom || "미지정",
      to: newTo,
      message: newMessage,
      category: matched.category,
      deadline: newDeadline || "미정",
      status: "pending",
      autoLevel: matched.autoLevel,
      guide: matched.guide,
      channel: "수동 입력",
      timestamp: new Date().toLocaleString("ko-KR"),
    };
    setTasks([task, ...tasks]);
    setNewMessage("");
    setNewFrom("");
    setNewDeadline("");
    // GitHub에도 저장
    fetch("/api/tasks", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(task),
    }).catch(() => {});
  };

  const updateStatus = (id: string, status: TaskStatus) => {
    setTasks(tasks.map((t) => (t.id === id ? { ...t, status } : t)));
  };

  const removeTask = (id: string) => {
    setTasks(tasks.filter((t) => t.id !== id));
  };

  const runTask = async (id: string) => {
    const task = tasks.find((t) => t.id === id);
    if (!task) return;
    const matched = matchCategory(task.message);
    const steps: ExecutionStep[] = matched.steps.map((label) => ({ label, status: "pending" }));

    // 1. 실행 상태로 전환 + 첫 번째 스텝 running
    const initialSteps = steps.map((s, i) => ({
      ...s,
      status: i === 0 ? "running" as const : "pending" as const,
    }));
    setTasks((prev) =>
      prev.map((t) => t.id === id ? { ...t, status: "running" as TaskStatus, executionSteps: initialSteps } : t)
    );

    try {
      // 2. 실제 API 호출
      const res = await fetch("/api/agents", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ category: task.category }),
      });
      const result = await res.json();

      if (result.success) {
        // 3. 성공: 모든 스텝 done + 결과물 경로
        const doneSteps = steps.map((s) => ({ ...s, status: "done" as const }));
        setTasks((prev) =>
          prev.map((t) =>
            t.id === id
              ? { ...t, status: "done" as TaskStatus, executionSteps: doneSteps, outputFile: result.outputPath }
              : t
          )
        );
      } else {
        // 4. 실패: 에러 표시
        const errorSteps = steps.map((s, i) => ({
          ...s,
          status: i === 0 ? "error" as const : "pending" as const,
        }));
        // 에러 상세 정보를 마지막 스텝으로 추가
        errorSteps.push({
          label: `오류: ${result.stderr?.slice(0, 200) || result.error || "알 수 없는 오류"}`,
          status: "error" as const,
        });
        setTasks((prev) =>
          prev.map((t) =>
            t.id === id
              ? { ...t, status: "in_progress" as TaskStatus, executionSteps: errorSteps }
              : t
          )
        );
      }
    } catch (err) {
      // 5. 네트워크 오류
      const errorSteps: ExecutionStep[] = [
        { label: `서버 연결 실패: ${err}`, status: "error" as const },
      ];
      setTasks((prev) =>
        prev.map((t) =>
          t.id === id
            ? { ...t, status: "in_progress" as TaskStatus, executionSteps: errorSteps }
            : t
        )
      );
    }
  };

  const handleMeetingUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const meeting: MeetingSummary = {
      id: Date.now().toString(),
      fileName: file.name,
      date: new Date().toISOString().split("T")[0],
      summary: "",
      participants: [],
      taskAssignments: [],
      directionChanges: [],
      status: "processing",
    };
    setMeetings([meeting, ...meetings]);

    // 시뮬레이션: 실제로는 API가 summarize.py를 호출
    setTimeout(() => {
      setMeetings((prev) =>
        prev.map((m) =>
          m.id === meeting.id
            ? {
                ...m,
                status: "done",
                summary: `${file.name} 회의록 요약 완료. 구조화된 요약, 업무 지시, 방향성 변화가 추출되었습니다.`,
                participants: ["이창준", "주호연", "임성욱"],
                taskAssignments: [
                  { assignee: "주호연", task: "Placement 26Q1 초안 작성", deadline: "2026-03-28" },
                  { assignee: "임성욱", task: "매크로 3월 업데이트", deadline: "2026-03-31" },
                ],
                directionChanges: [
                  { fromWho: "이창준", content: "Claude Code 위에서 git push/pull로 협업하는 방향으로 전환" },
                ],
                notionUrl: "https://notion.so/meeting-note-example",
              }
            : m
        )
      );
    }, 3000);

    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const filtered = filter === "all" ? tasks : tasks.filter((t) => t.status === filter);
  const stats = {
    total: tasks.length,
    pending: tasks.filter((t) => t.status === "pending").length,
    active: tasks.filter((t) => t.status === "in_progress" || t.status === "running").length,
    done: tasks.filter((t) => t.status === "done").length,
  };

  return (
    <div className="min-h-screen p-6 max-w-6xl mx-auto">
      {/* 헤더 */}
      <header className="mb-8">
        <h1 className="text-2xl font-bold">
          <span style={{ color: "var(--accent)" }}>전략추진실</span> 워크스페이스
        </h1>
        <p className="text-sm mt-1" style={{ color: "var(--text-muted)" }}>
          업무 자동화 대시보드 — 슬랙 기반 업무 관리 · AI 에이전트 실행
        </p>
      </header>

      {/* 통계 */}
      <div className="grid grid-cols-4 gap-4 mb-8">
        {[
          { label: "전체", value: stats.total, color: "var(--text)" },
          { label: "대기", value: stats.pending, color: "var(--yellow)" },
          { label: "진행 중", value: stats.active, color: "var(--blue)" },
          { label: "완료", value: stats.done, color: "var(--accent)" },
        ].map((s) => (
          <div key={s.label} className="p-4 rounded-xl" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
            <div className="text-xs" style={{ color: "var(--text-muted)" }}>{s.label}</div>
            <div className="text-2xl font-bold mt-1" style={{ color: s.color }}>{s.value}</div>
          </div>
        ))}
      </div>

      {/* 탭 네비게이션 */}
      <div className="flex gap-1 mb-6 p-1 rounded-xl" style={{ background: "var(--surface)" }}>
        {([
          { key: "tasks", label: "📋 업무 관리", count: stats.total },
          { key: "meeting", label: "🎙️ 회의록", count: meetings.length },
          { key: "knowledge", label: "📚 지식 베이스", count: KNOWLEDGE_BASE.length },
        ] as { key: TabType; label: string; count: number }[]).map((tab) => (
          <button key={tab.key} onClick={() => setActiveTab(tab.key)}
            className={`flex-1 px-4 py-2 rounded-lg text-sm font-medium transition-all ${activeTab === tab.key ? "text-black" : "text-gray-400 hover:text-gray-300"}`}
            style={{ background: activeTab === tab.key ? "var(--accent)" : "transparent" }}>
            {tab.label} ({tab.count})
          </button>
        ))}
      </div>

      {/* ─── 탭: 업무 관리 ─── */}
      {activeTab === "tasks" && (
        <>
          {/* 업무 입력 */}
          <div className="p-4 rounded-xl mb-6" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
            <h2 className="text-sm font-medium mb-3" style={{ color: "var(--text-muted)" }}>📨 업무 추가</h2>
            <div className="flex gap-3 mb-3">
              <input value={newFrom} onChange={(e) => setNewFrom(e.target.value)}
                placeholder="보내는 사람"
                className="w-36 px-3 py-2 rounded-lg text-sm bg-black border border-gray-700 text-white placeholder-gray-500" />
              <span className="py-2 text-gray-500">→</span>
              <input value={newTo} onChange={(e) => setNewTo(e.target.value)}
                placeholder="받는 사람"
                className="w-36 px-3 py-2 rounded-lg text-sm bg-black border border-gray-700 text-white placeholder-gray-500" />
              <input type="date" value={newDeadline} onChange={(e) => setNewDeadline(e.target.value)}
                className="px-3 py-2 rounded-lg text-sm bg-black border border-gray-700 text-white" />
            </div>
            <div className="flex gap-3">
              <input
                value={newMessage}
                onChange={(e) => setNewMessage(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && addTask()}
                placeholder="슬랙 메시지 내용 입력... (예: 26Q1 해줘, 매크로 업데이트, 이거 번역해줘)"
                className="flex-1 px-4 py-2 rounded-lg text-sm bg-black border border-gray-700 text-white placeholder-gray-500"
              />
              <button onClick={addTask}
                className="px-4 py-2 rounded-lg text-sm font-medium text-black"
                style={{ background: "var(--accent)" }}>
                추가
              </button>
            </div>
          </div>

          {/* 필터 */}
          <div className="flex gap-2 mb-4">
            {(["all", "pending", "in_progress", "running", "done"] as const).map((f) => (
              <button key={f} onClick={() => setFilter(f)}
                className={`px-3 py-1 rounded-lg text-xs transition-all ${filter === f ? "text-black" : "text-gray-400"}`}
                style={{ background: filter === f ? "var(--accent)" : "var(--surface)", border: "1px solid var(--border)" }}>
                {{ all: "전체", pending: "대기", in_progress: "진행 중", running: "실행 중", done: "완료" }[f]}
              </button>
            ))}
          </div>

          {/* 업무 대기목록 */}
          <div className="space-y-3 mb-8">
            {filtered.length === 0 && (
              <div className="text-center py-12 text-sm" style={{ color: "var(--text-muted)" }}>
                업무가 없습니다. 위에서 추가하거나 슬랙 연동을 설정하세요.
              </div>
            )}
            {filtered.map((task) => (
              <div key={task.id} className="p-4 rounded-xl transition-all" style={{ background: "var(--surface)", border: `1px solid ${task.status === "running" ? "var(--accent)" : "var(--border)"}` }}>
                <div className="flex items-start justify-between mb-2">
                  <div className="flex items-center gap-2">
                    <StatusBadge status={task.status} />
                    <AutoBadge level={task.autoLevel} />
                    <span className="text-sm font-medium">{task.category}</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs" style={{ color: "var(--text-muted)" }}>~{task.deadline}</span>
                    {task.status === "done" && (
                      <button onClick={() => removeTask(task.id)} className="text-xs text-red-400 hover:text-red-300">✕</button>
                    )}
                  </div>
                </div>

                <p className="text-sm mb-2">&ldquo;{task.message}&rdquo;</p>

                <div className="flex items-center justify-between">
                  <div className="text-xs" style={{ color: "var(--text-muted)" }}>
                    {task.from} → {task.to} · {task.channel} · {task.timestamp}
                  </div>
                  <div className="flex gap-2">
                    {task.status === "pending" && (
                      <button onClick={() => updateStatus(task.id, "in_progress")}
                        className="px-3 py-1 rounded text-xs text-blue-400 border border-blue-500/30 hover:bg-blue-500/10 transition-all">
                        시작
                      </button>
                    )}
                    {task.status === "in_progress" && (
                      <button onClick={() => updateStatus(task.id, "done")}
                        className="px-3 py-1 rounded text-xs text-green-400 border border-green-500/30 hover:bg-green-500/10 transition-all">
                        완료
                      </button>
                    )}
                    {task.autoLevel === "auto" && (task.status === "pending" || task.status === "in_progress") && (
                      <button onClick={() => runTask(task.id)}
                        className="px-3 py-1 rounded text-xs text-black font-medium transition-all hover:opacity-80"
                        style={{ background: "var(--accent)" }}>
                        ▶ 실행
                      </button>
                    )}
                  </div>
                </div>

                {task.executionSteps && task.executionSteps.length > 0 ? (
                  <ExecutionLog steps={task.executionSteps} outputFile={task.outputFile} />
                ) : (
                  <TaskGuide task={task} />
                )}
              </div>
            ))}
          </div>

          {/* 자동 처리 Skill 현황 */}
          <h2 className="text-sm font-medium mb-3" style={{ color: "var(--text-muted)" }}>⚡ 에이전트 현황 (6개 중 4개 구현)</h2>
          <div className="grid grid-cols-3 gap-3">
            <SkillCard name="예산·구매 품의" ready={true} command="node fill.js" />
            <SkillCard name="장표 번역" ready={true} command="python translate.py" />
            <SkillCard name="Macro 분석" ready={true} command="python update_macro.py" />
            <SkillCard name="회의록 생성" ready={true} command="python summarize.py" />
            <SkillCard name="장표 제작" ready={false} command="Claude/Genspark (미구현)" />
            <SkillCard name="Placement 분석" ready={false} command="run_jk → calc_rms → gen_ppt (미구현)" />
          </div>
        </>
      )}

      {/* ─── 탭: 회의록 ─── */}
      {activeTab === "meeting" && (
        <div>
          {/* 업로드 영역 */}
          <div className="p-6 rounded-xl mb-6 text-center" style={{ background: "var(--surface)", border: "2px dashed var(--border)" }}>
            <div className="text-3xl mb-2">🎙️</div>
            <p className="text-sm mb-3">회의록 TXT 파일을 업로드하면 자동으로 요약합니다</p>
            <p className="text-xs mb-4" style={{ color: "var(--text-muted)" }}>
              업무 지시 → 대시보드 업무 추가 · 방향성 변화 → Team Context 업데이트
            </p>
            <input ref={fileInputRef} type="file" accept=".txt" onChange={handleMeetingUpload} className="hidden" />
            <button onClick={() => fileInputRef.current?.click()}
              className="px-6 py-2 rounded-lg text-sm font-medium text-black"
              style={{ background: "var(--accent)" }}>
              📂 TXT 파일 선택
            </button>
          </div>

          {/* 회의록 목록 */}
          <div className="space-y-4">
            {meetings.length === 0 && (
              <div className="text-center py-12 text-sm" style={{ color: "var(--text-muted)" }}>
                아직 정리된 회의록이 없습니다. TXT 파일을 업로드해보세요.
              </div>
            )}
            {meetings.map((m) => (
              <div key={m.id} className="p-4 rounded-xl" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center gap-2">
                    <span className={`px-2 py-0.5 text-xs rounded border ${m.status === "done" ? "bg-green-500/20 text-green-400 border-green-500/30" : m.status === "processing" ? "bg-purple-500/20 text-purple-400 border-purple-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>
                      {m.status === "done" ? "완료" : m.status === "processing" ? "처리 중..." : "오류"}
                    </span>
                    <span className="text-sm font-medium">{m.fileName}</span>
                  </div>
                  <span className="text-xs" style={{ color: "var(--text-muted)" }}>{m.date}</span>
                </div>

                {m.status === "processing" && (
                  <div className="flex items-center gap-2 text-xs text-purple-400">
                    <span className="animate-pulse">⏳</span> Claude가 요약 중...
                  </div>
                )}

                {m.status === "done" && (
                  <div className="space-y-3">
                    <p className="text-sm">{m.summary}</p>
                    <div className="text-xs" style={{ color: "var(--text-muted)" }}>
                      참석자: {m.participants.join(", ")}
                    </div>

                    {m.taskAssignments.length > 0 && (
                      <div className="p-3 rounded-lg" style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
                        <div className="text-xs font-medium mb-2" style={{ color: "var(--accent)" }}>📋 추출된 업무 지시</div>
                        {m.taskAssignments.map((ta, i) => (
                          <div key={i} className="text-xs flex items-center justify-between py-1">
                            <span>→ {ta.assignee}: {ta.task}</span>
                            <span style={{ color: "var(--text-muted)" }}>~{ta.deadline}</span>
                          </div>
                        ))}
                      </div>
                    )}

                    {m.directionChanges.length > 0 && (
                      <div className="p-3 rounded-lg" style={{ background: "var(--bg)", border: "1px solid #2a1a3a" }}>
                        <div className="text-xs font-medium mb-2 text-purple-400">🧭 방향성 변화</div>
                        {m.directionChanges.map((dc, i) => (
                          <div key={i} className="text-xs py-1">
                            <span className="text-purple-300">{dc.fromWho}:</span> {dc.content}
                          </div>
                        ))}
                      </div>
                    )}

                    <div className="flex gap-2 mt-2">
                      {m.notionUrl && (
                        <a href={m.notionUrl} target="_blank" rel="noopener noreferrer"
                          className="px-3 py-1 rounded text-xs border border-gray-600 text-gray-300 hover:bg-gray-800 transition-all">
                          📄 Notion에서 보기
                        </a>
                      )}
                      <button className="px-3 py-1 rounded text-xs text-black font-medium"
                        style={{ background: "var(--accent)" }}
                        onClick={() => {
                          m.taskAssignments.forEach((ta) => {
                            const matched = matchCategory(ta.task);
                            const task: Task = {
                              id: Date.now().toString() + Math.random(),
                              from: "회의록",
                              to: ta.assignee,
                              message: ta.task,
                              category: matched.category,
                              deadline: ta.deadline,
                              status: "pending",
                              autoLevel: matched.autoLevel,
                              guide: matched.guide,
                              channel: `회의록: ${m.fileName}`,
                              timestamp: m.date,
                            };
                            setTasks((prev) => [task, ...prev]);
                          });
                          setActiveTab("tasks");
                        }}>
                        📨 업무 대기목록에 추가
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ─── 탭: 지식 베이스 ─── */}
      {activeTab === "knowledge" && (
        <div>
          <p className="text-sm mb-4" style={{ color: "var(--text-muted)" }}>
            자동화되지 않은 업무의 진행 절차와 담당자 정보입니다. 해당 업무가 슬랙으로 들어오면 여기서 가이드를 확인하세요.
          </p>
          <div className="space-y-3">
            {KNOWLEDGE_BASE.map((item, i) => (
              <KnowledgeCard key={i} item={item} />
            ))}
          </div>
        </div>
      )}

      {/* 푸터 */}
      <footer className="text-center text-xs py-8 mt-8" style={{ color: "var(--text-muted)", borderTop: "1px solid var(--border)" }}>
        전략추진실 인턴 워크스페이스 v1.0 — Claude Code 기반 · worxphere-auto
      </footer>
    </div>
  );
}
