"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";

type TaskStatus = "pending" | "in_progress" | "done";
type AutoLevel = "auto" | "manual" | "knowledge";

interface Task {
  id: string;
  from: string;
  to: string;
  message: string;
  category: string;
  startDate: string;
  deadline: string;
  status: TaskStatus;
  autoLevel: AutoLevel;
  guide: string;
  channel: string;
  timestamp: string;
  notes?: string[];
  steps?: string[];
  topicFile?: string;
}

interface MeetingNote {
  id: string;
  title: string;
  date: string;
  decisions: string;
  url: string;
}

// 에이전트 실행 모드 topic (autoLevel=auto + 실행 가능한 스크립트 있음)
const AGENT_TOPICS = new Set([
  "regular/macro-update",
  "regular/placement-analysis",
  "fluid/ppt-create",
  "fluid/ppt-translate",
]);

// topic별 담당자 (Context/Directory.md 기반)
const TOPIC_CONTACTS: Record<string, string[]> = {
  "regular/macro-update":       ["창준님"],
  "regular/macro-indicators":   ["창준님"],
  "regular/placement-update":   ["엠브레인 문주원님", "창준님"],
  "regular/placement-analysis": ["엠브레인 문주원님", "창준님"],
  "fluid/ppt-create":           ["창준님"],
  "fluid/ppt-translate":        ["창준님"],
  "fluid/academia-contract":    ["총무팀 이민희님", "총무팀 남영현님", "KT alpha 강석현님"],
  "budget/placement-concur":    ["엠브레인 문주원님", "재무회계팀 박은미님/왕윤형님"],
  "budget/enkoline-concur":     ["재무회계팀 박은미님/왕윤형님"],
  "budget/consulting-concur":   ["재무회계팀 박은미님/왕윤형님"],
  "budget/law-firm":            ["법무팀 윤종화님", "법무팀 박민경님", "재무회계팀"],
  "budget/ninehire":            ["창준님", "재무회계팀"],
  "budget/gifticon":            ["KT alpha 강석현님", "ATS운영팀 최돈민님", "창준님"],
  "budget/budget-transfer":     ["경영기획팀 임장식님", "창준님"],
  "budget/vendor-registration": ["재무회계팀 박은미님/왕윤형님"],
  "budget/budget-101":          ["경영기획팀 임장식님"],
};

function calcDday(deadline: string): string | null {
  if (!deadline || deadline === "미정") return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const due = new Date(deadline);
  due.setHours(0, 0, 0, 0);
  const diff = Math.round((due.getTime() - today.getTime()) / 86400000);
  if (diff === 0) return "D-day";
  if (diff > 0) return `D-${diff}`;
  return `D+${Math.abs(diff)}`;
}

export default function TaskDetailPage() {
  const params = useParams();
  const router = useRouter();
  const [task, setTask] = useState<Task | null>(null);
  const [loading, setLoading] = useState(true);
  const [context, setContext] = useState<MeetingNote[]>([]);
  const [contextLoading, setContextLoading] = useState(false);
  const [agentRunning, setAgentRunning] = useState(false);
  const [agentResult, setAgentResult] = useState<string | null>(null);
  const [expandedNote, setExpandedNote] = useState<string | null>(null);
  const [placementQuarter, setPlacementQuarter] = useState("26Q1");

  useEffect(() => {
    fetch("/api/tasks")
      .then((r) => r.json())
      .then((data) => {
        const found = data.tasks?.find((t: { id: string }) => t.id === params.id);
        if (found) {
          setTask({ ...found, startDate: found.startDate || "" });
          if (found.topicFile) {
            setContextLoading(true);
            fetch(`/api/context?topicFile=${encodeURIComponent(found.topicFile)}`)
              .then((r) => r.json())
              .then((d) => { setContext(d.results || []); setContextLoading(false); })
              .catch(() => setContextLoading(false));
          }
        }
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, [params.id]);

  const updateStatus = (status: TaskStatus) => {
    if (!task) return;
    setTask({ ...task, status });
    fetch("/api/tasks", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: task.id, status }),
    }).catch(() => {});
  };

  const runAgent = async () => {
    if (!task) return;
    setAgentRunning(true);
    setAgentResult(null);
    const agentParams: Record<string, string> | undefined =
      task.topicFile === "regular/placement-analysis"
        ? { quarter: placementQuarter }
        : undefined;
    try {
      const res = await fetch("/api/agents", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ topicFile: task.topicFile, params: agentParams }),
      });
      const data = await res.json();
      setAgentResult(data.success
        ? `완료${data.outputPath ? ` → ${data.outputPath}` : ""}`
        : `오류: ${(data.stderr || data.error || "").slice(0, 200)}`
      );
    } catch (e) {
      setAgentResult(`서버 연결 실패: ${e}`);
    } finally {
      setAgentRunning(false);
    }
  };

  if (loading) return (
    <div className="min-h-screen flex items-center justify-center" style={{ color: "var(--text-muted)" }}>
      불러오는 중...
    </div>
  );

  if (!task) return (
    <div className="min-h-screen flex flex-col items-center justify-center gap-4">
      <div className="text-sm" style={{ color: "var(--text-muted)" }}>업무를 찾을 수 없습니다.</div>
      <button onClick={() => router.push("/")} className="px-4 py-2 rounded-lg text-sm text-black font-medium" style={{ background: "var(--accent)" }}>
        ← 메인으로
      </button>
    </div>
  );

  const dday = calcDday(task.deadline);
  const ddayColor = !dday ? "var(--text-muted)" : dday === "D-day" ? "#f59e0b" : dday.startsWith("D+") ? "#ef4444" : parseInt(dday.replace("D-", "")) <= 3 ? "#f97316" : "#9ca3af";
  const isAgentMode = AGENT_TOPICS.has(task.topicFile ?? "");
  const contacts = TOPIC_CONTACTS[task.topicFile ?? ""] ?? [];

  return (
    <div className="min-h-screen p-6 max-w-3xl mx-auto">
      {/* 뒤로가기 */}
      <button onClick={() => router.push("/")}
        className="flex items-center gap-1 text-sm mb-6 hover:opacity-80 transition-all"
        style={{ color: "var(--text-muted)" }}>
        ← 업무 목록
      </button>

      {/* 헤더 */}
      <div className="p-5 rounded-2xl mb-4" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
        <div className="flex items-start justify-between mb-3">
          <div className="flex items-center gap-2 flex-wrap">
            <span className={`px-2 py-0.5 text-xs rounded border ${
              task.status === "pending" ? "bg-yellow-500/20 text-yellow-400 border-yellow-500/30" :
              task.status === "in_progress" ? "bg-blue-500/20 text-blue-400 border-blue-500/30" :
              "bg-green-500/20 text-green-400 border-green-500/30"
            }`}>
              {task.status === "pending" ? "대기" : task.status === "in_progress" ? "진행 중" : "완료"}
            </span>
            <span className={`px-2 py-0.5 text-xs rounded border ${
              isAgentMode ? "bg-green-500/20 text-green-400 border-green-500/30" :
              task.autoLevel === "manual" ? "bg-yellow-500/20 text-yellow-400 border-yellow-500/30" :
              "bg-gray-500/20 text-gray-400 border-gray-500/30"
            }`}>
              {isAgentMode ? "🟢 에이전트" : task.autoLevel === "manual" ? "🟡 가이드" : "📚 지식"}
            </span>
            <h1 className="text-lg font-bold">{task.category}</h1>
          </div>
          {dday && (
            <span className="text-sm font-bold px-2 py-1 rounded shrink-0"
              style={{ color: ddayColor, background: `${ddayColor}22` }}>
              {dday}
            </span>
          )}
        </div>

        <p className="text-sm mb-3 text-gray-200">&ldquo;{task.message}&rdquo;</p>

        <div className="flex flex-wrap gap-4 text-xs" style={{ color: "var(--text-muted)" }}>
          <span>지시자: <strong className="text-white">{task.from}</strong></span>
          <span>수행자: <strong className="text-white">{task.to}</strong></span>
          {task.startDate && <span>시작: <strong className="text-white">{task.startDate}</strong></span>}
          <span>마감: <strong className="text-white">{task.deadline}</strong></span>
          <span>{task.channel} · {task.timestamp}</span>
        </div>
      </div>

      {/* ── 에이전트 모드 ── */}
      {isAgentMode && (
        <div className="p-5 rounded-2xl mb-4" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <h2 className="text-sm font-semibold mb-1" style={{ color: "var(--accent)" }}>에이전트 실행</h2>
          <p className="text-xs mb-4" style={{ color: "var(--text-muted)" }}>{task.guide}</p>

          {task.steps && task.steps.length > 0 && (
            <div className="space-y-2 mb-4">
              {task.steps.map((step, i) => (
                <div key={i} className="flex items-center gap-3 p-2.5 rounded-xl text-xs"
                  style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
                  <span className="w-5 h-5 rounded-full flex items-center justify-center text-xs font-bold shrink-0 text-black"
                    style={{ background: "var(--accent)" }}>{i + 1}</span>
                  <span>{step}</span>
                </div>
              ))}
            </div>
          )}

          {task.topicFile === "regular/placement-analysis" && (
            <div className="flex items-center gap-2 mb-3">
              <label className="text-xs shrink-0" style={{ color: "var(--text-muted)" }}>분기</label>
              <input
                value={placementQuarter}
                onChange={(e) => setPlacementQuarter(e.target.value)}
                placeholder="예: 26Q1"
                className="flex-1 px-3 py-1.5 rounded-lg text-xs"
                style={{ background: "var(--bg)", border: "1px solid var(--border)", color: "var(--text)" }}
              />
            </div>
          )}

          <button onClick={runAgent} disabled={agentRunning}
            className={`w-full py-3 rounded-xl text-sm font-semibold transition-all ${
              agentRunning ? "bg-gray-700 text-gray-400 cursor-wait" : "text-black hover:opacity-80"
            }`}
            style={agentRunning ? {} : { background: "var(--accent)" }}>
            {agentRunning ? "⏳ 에이전트 실행 중..." : "▶ 에이전트 실행"}
          </button>

          {agentResult && (
            <div className={`mt-3 p-3 rounded-xl text-xs ${
              agentResult.startsWith("오류") || agentResult.startsWith("서버")
                ? "bg-red-500/10 text-red-400" : "bg-green-500/10 text-green-400"
            }`}>
              {agentResult}
            </div>
          )}
        </div>
      )}

      {/* ── 가이드 모드 ── */}
      {!isAgentMode && (
        <div className="p-5 rounded-2xl mb-4" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <h2 className="text-sm font-semibold mb-1" style={{ color: "var(--accent)" }}>작업 방식</h2>
          <p className="text-xs mb-4" style={{ color: "var(--text-muted)" }}>{task.guide}</p>

          {task.steps && task.steps.length > 0 && (
            <div className="space-y-2 mb-4">
              {task.steps.map((step, i) => (
                <div key={i} className="flex items-start gap-3 p-3 rounded-xl text-sm"
                  style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
                  <span className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold shrink-0 text-black"
                    style={{ background: "var(--accent)" }}>{i + 1}</span>
                  <span>{step}</span>
                </div>
              ))}
            </div>
          )}

          {contacts.length > 0 && (
            <div className="p-3 rounded-xl" style={{ background: "var(--bg)", border: "1px solid var(--border)" }}>
              <div className="text-xs mb-2" style={{ color: "var(--text-muted)" }}>담당자</div>
              <div className="flex flex-wrap gap-2">
                {contacts.map((c, i) => (
                  <span key={i} className="px-2 py-1 text-xs rounded-lg bg-gray-800 text-gray-300 border border-gray-700">{c}</span>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── 과거 맥락 (Notion Meeting Notes) ── */}
      <div className="p-5 rounded-2xl mb-4" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
        <h2 className="text-sm font-semibold mb-1" style={{ color: "var(--accent)" }}>과거 맥락</h2>
        <p className="text-xs mb-3" style={{ color: "var(--text-muted)" }}>관련 싱크 기록</p>

        {contextLoading ? (
          <div className="p-3 text-xs text-center text-gray-500">로딩 중...</div>
        ) : context.length > 0 ? (
          <div className="space-y-2">
            {context.map((note) => (
              <div key={note.id} className="rounded-xl overflow-hidden"
                style={{ border: "1px solid var(--border)" }}>
                <div className="p-3 flex items-center justify-between cursor-pointer hover:bg-white/5 transition-all"
                  style={{ background: "var(--bg)" }}
                  onClick={() => setExpandedNote(expandedNote === note.id ? null : note.id)}>
                  <div className="flex items-center gap-2 min-w-0">
                    <span className="text-xs shrink-0" style={{ color: "var(--text-muted)" }}>{note.date}</span>
                    <span className="text-sm truncate">{note.title}</span>
                  </div>
                  <span className="text-xs text-gray-500 shrink-0 ml-2">{expandedNote === note.id ? "▲" : "▼"}</span>
                </div>
                {expandedNote === note.id && (
                  <div className="p-3 border-t text-xs text-gray-300 whitespace-pre-wrap leading-relaxed"
                    style={{ borderColor: "var(--border)" }}>
                    {note.decisions}
                  </div>
                )}
              </div>
            ))}
          </div>
        ) : (
          <div className="p-3 rounded-xl text-xs text-gray-500 text-center"
            style={{ border: "1px dashed var(--border)" }}>
            관련 싱크 기록이 없습니다
          </div>
        )}
      </div>

      {/* ── 현재 맥락 (슬랙 스레드) ── */}
      {task.notes && task.notes.length > 0 && (
        <div className="p-5 rounded-2xl mb-6" style={{ background: "var(--surface)", border: "1px solid var(--border)" }}>
          <h2 className="text-sm font-semibold mb-1" style={{ color: "var(--accent)" }}>현재 맥락</h2>
          <p className="text-xs mb-3" style={{ color: "var(--text-muted)" }}>슬랙 스레드 추가 정보</p>
          <div className="space-y-1">
            {task.notes.map((note, i) => (
              <div key={i} className="p-2 rounded text-xs text-gray-300" style={{ background: "var(--bg)" }}>
                · {note}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── 액션 버튼 ── */}
      <div className="flex gap-3">
        <button onClick={() => { updateStatus("pending"); router.push("/"); }}
          className="px-4 py-2 rounded-lg text-sm border border-gray-700 text-gray-400 hover:bg-gray-800 transition-all">
          ↩ 대기로 되돌리기
        </button>
        <button onClick={() => { updateStatus("done"); router.push("/"); }}
          className="flex-1 px-4 py-2 rounded-lg text-sm font-medium text-black transition-all hover:opacity-80"
          style={{ background: "var(--accent)" }}>
          완료 처리 →
        </button>
      </div>
    </div>
  );
}
