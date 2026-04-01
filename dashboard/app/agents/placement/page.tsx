"use client";

import { useState, useRef } from "react";
import Link from "next/link";

interface Step {
  label: string;
  description: string;
  status: "locked" | "ready" | "running" | "done" | "error";
  output?: string;
}

export default function PlacementAgent() {
  const [quarter, setQuarter] = useState("");
  const [jkFile, setJkFile] = useState<File | null>(null);
  const [amFile, setAmFile] = useState<File | null>(null);
  const [inputConfirmed, setInputConfirmed] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [outputPath, setOutputPath] = useState("");
  const [jkPath, setJkPath] = useState("");
  const [amPath, setAmPath] = useState("");
  const jkRef = useRef<HTMLInputElement>(null);
  const amRef = useRef<HTMLInputElement>(null);

  const [steps, setSteps] = useState<Step[]>([
    { label: "JK Raw 데이터 분류", description: "JobKorea Raw 데이터를 분류표에 매칭하여 R_통합 시트를 생성합니다.", status: "locked" },
    { label: "AM Raw 데이터 분류", description: "AlbaMain Raw 데이터를 분류표에 매칭하여 R_통합 시트를 생성합니다.", status: "locked" },
    { label: "RMS 계산", description: "JK/AM 각각 14개 시트 RMS를 계산합니다.", status: "locked" },
    { label: "PPT 자동 생성", description: "분기 PPT를 자동 생성합니다 (6슬라이드).", status: "locked" },
    { label: "수작업 확인", description: "산점도/부록 슬라이드는 수동 보완이 필요합니다.", status: "locked" },
  ]);
  const [running, setRunning] = useState(false);

  const confirmInputs = async () => {
    if (!quarter.match(/^\d{2}Q[1-4]$/i) || !jkFile || !amFile) return;
    setUploading(true);

    try {
      const uploadOne = async (file: File) => {
        const form = new FormData();
        form.append("file", file);
        const res = await fetch("/api/upload", { method: "POST", body: form });
        return res.json();
      };

      const [jkRes, amRes] = await Promise.all([uploadOne(jkFile), uploadOne(amFile)]);

      if (jkRes.success && amRes.success) {
        setJkPath(jkRes.path);
        setAmPath(amRes.path);
        setInputConfirmed(true);
        setSteps((prev) => prev.map((s, i) => (i === 0 ? { ...s, status: "ready" } : s)));
      }
    } catch (err) {
      console.error(err);
    }
    setUploading(false);
  };

  const runStep = async (stepIdx: number) => {
    setRunning(true);
    setSteps((prev) => prev.map((s, i) => (i === stepIdx ? { ...s, status: "running", output: undefined } : s)));

    // 마지막 단계(수작업)는 확인만
    if (stepIdx === 4) {
      await new Promise((r) => setTimeout(r, 500));
      setSteps((prev) => prev.map((s, i) => (i === 4 ? { ...s, status: "done", output: "산점도(슬라이드 7-10, 14-17)와 부록을 수동으로 보완해주세요." } : s)));
      setRunning(false);
      return;
    }

    try {
      const res = await fetch("/api/agents/placement", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ jkPath, amPath, quarter: quarter.toUpperCase(), step: stepIdx + 1 }),
      });
      const data = await res.json();

      if (data.success) {
        if (data.outputPath) setOutputPath(data.outputPath);
        setSteps((prev) => prev.map((s, i) => {
          if (i === stepIdx) return { ...s, status: "done", output: data.output };
          if (i === stepIdx + 1 && s.status === "locked") return { ...s, status: "ready" };
          return s;
        }));
      } else {
        setSteps((prev) => prev.map((s, i) => (i === stepIdx ? { ...s, status: "error", output: data.error || "실패" } : s)));
      }
    } catch (err) {
      setSteps((prev) => prev.map((s, i) => (i === stepIdx ? { ...s, status: "error", output: `서버 연결 실패: ${err}` } : s)));
    }
    setRunning(false);
  };

  const allDone = steps.every((s) => s.status === "done");

  return (
    <div className="min-h-screen p-6 max-w-4xl mx-auto">
      <Link href="/" className="text-sm text-gray-400 hover:text-white transition-all">← 대시보드</Link>

      <div className="mt-6 mb-8">
        <div className="flex items-center gap-3 mb-2">
          <h1 className="text-2xl font-bold">📈 Placement Survey 분석</h1>
          <span className="text-xs px-2 py-0.5 rounded-full bg-green-500/20 text-green-400">● 활성</span>
        </div>
        <p className="text-sm" style={{ color: "var(--text-muted)" }}>
          JK/AM Raw 데이터 → R_통합 분류 → RMS 계산 → PPT 자동 생성
        </p>
      </div>

      {/* 입력 */}
      <div className={`p-5 rounded-xl mb-6 border ${inputConfirmed ? "border-green-500/30 bg-green-500/5" : "border-gray-600"}`} style={inputConfirmed ? {} : { background: "var(--surface)" }}>
        <h2 className="text-sm font-medium mb-3" style={{ color: "var(--text-muted)" }}>입력 데이터</h2>

        <div className="space-y-3">
          <div>
            <label className="text-xs text-gray-400 block mb-1">분기</label>
            <div className="flex gap-2">
              <input value={quarter} onChange={(e) => setQuarter(e.target.value)} placeholder="예: 26Q1" disabled={inputConfirmed}
                className="flex-1 px-4 py-2.5 rounded-lg text-sm bg-black border border-gray-700 text-white placeholder-gray-500 disabled:opacity-50" />
              {["25Q4", "26Q1", "26Q2"].map((q) => (
                <button key={q} onClick={() => setQuarter(q)} disabled={inputConfirmed}
                  className="px-3 py-2 rounded-lg text-xs bg-gray-800 text-gray-400 hover:text-white disabled:opacity-30">{q}</button>
              ))}
            </div>
          </div>

          <div>
            <label className="text-xs text-gray-400 block mb-1">JK Raw 데이터 (xlsx)</label>
            <input ref={jkRef} type="file" accept=".xlsx,.xls" className="hidden"
              onChange={(e) => setJkFile(e.target.files?.[0] || null)} />
            <button onClick={() => jkRef.current?.click()} disabled={inputConfirmed}
              className="w-full px-4 py-2.5 rounded-lg text-sm text-left bg-black border border-gray-700 text-white disabled:opacity-50 hover:border-gray-500">
              {jkFile ? `📄 ${jkFile.name}` : "JK Raw 파일 선택..."}
            </button>
          </div>

          <div>
            <label className="text-xs text-gray-400 block mb-1">AM Raw 데이터 (xlsx)</label>
            <input ref={amRef} type="file" accept=".xlsx,.xls" className="hidden"
              onChange={(e) => setAmFile(e.target.files?.[0] || null)} />
            <button onClick={() => amRef.current?.click()} disabled={inputConfirmed}
              className="w-full px-4 py-2.5 rounded-lg text-sm text-left bg-black border border-gray-700 text-white disabled:opacity-50 hover:border-gray-500">
              {amFile ? `📄 ${amFile.name}` : "AM Raw 파일 선택..."}
            </button>
          </div>
        </div>

        {!inputConfirmed ? (
          <button onClick={confirmInputs} disabled={!quarter.match(/^\d{2}Q[1-4]$/i) || !jkFile || !amFile || uploading}
            className="mt-4 w-full px-6 py-2.5 rounded-lg text-sm font-medium text-black hover:opacity-80 disabled:opacity-30"
            style={{ background: "var(--accent)" }}>
            {uploading ? "⏳ 업로드 중..." : "확정 및 업로드"}
          </button>
        ) : (
          <div className="mt-3 text-sm text-green-400">✅ {quarter.toUpperCase()} — 파일 업로드 완료</div>
        )}
      </div>

      {/* 단계별 진행 */}
      <div className="space-y-4">
        {steps.map((step, i) => (
          <div key={i} className={`rounded-xl border transition-all ${
            step.status === "done" ? "border-green-500/30 bg-green-500/5" :
            step.status === "running" ? "border-yellow-500/30 bg-yellow-500/5" :
            step.status === "error" ? "border-red-500/30 bg-red-500/5" :
            step.status === "ready" ? "border-gray-600 bg-gray-800/30" :
            "border-gray-800 bg-gray-900/30 opacity-50"
          }`}>
            <div className="p-5">
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-3">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold ${
                    step.status === "done" ? "bg-green-500 text-black" :
                    step.status === "running" ? "bg-yellow-500 text-black animate-pulse" :
                    step.status === "error" ? "bg-red-500 text-white" :
                    step.status === "ready" ? "bg-gray-600 text-white" :
                    "bg-gray-800 text-gray-600"
                  }`}>
                    {step.status === "done" ? "✓" : step.status === "error" ? "✗" : i + 1}
                  </div>
                  <div>
                    <div className={`font-medium text-sm ${
                      step.status === "done" ? "text-green-400" :
                      step.status === "running" ? "text-yellow-400" :
                      step.status === "error" ? "text-red-400" :
                      step.status === "ready" ? "text-white" :
                      "text-gray-600"
                    }`}>{step.label}</div>
                    <div className="text-xs text-gray-500 mt-0.5">{step.description}</div>
                  </div>
                </div>
                {step.status === "ready" && (
                  <button onClick={() => runStep(i)} disabled={running}
                    className={`px-4 py-1.5 rounded-lg text-xs font-medium transition-all ${running ? "bg-gray-700 text-gray-500" : "text-black hover:opacity-80"}`}
                    style={running ? {} : { background: "var(--accent)" }}>
                    {running ? "⏳" : "▶ 실행"}
                  </button>
                )}
                {step.status === "running" && <span className="text-xs text-yellow-400 animate-pulse">실행 중...</span>}
              </div>
              {step.output && step.status !== "locked" && (
                <div className={`mt-3 p-3 rounded-lg text-xs font-mono whitespace-pre-wrap ${
                  step.status === "error" ? "bg-red-500/10 text-red-300" : "bg-black/30 text-gray-300"
                }`}>{step.output}</div>
              )}
            </div>
          </div>
        ))}
      </div>

      {allDone && (
        <div className="mt-6 p-5 rounded-xl bg-green-500/10 border border-green-500/30 text-center">
          <div className="text-green-400 font-medium mb-1">✅ Placement {quarter.toUpperCase()} 분석 완료</div>
          {outputPath && (
            <a href={`/api/download?path=${encodeURIComponent(outputPath)}`}
              className="inline-block mt-3 px-6 py-2.5 rounded-lg text-sm font-medium text-black hover:opacity-80"
              style={{ background: "var(--accent)" }}>
              📥 결과 PPT 다운로드
            </a>
          )}
          <p className="text-xs text-gray-400 mt-2">산점도/부록은 수동 보완해주세요.</p>
        </div>
      )}
    </div>
  );
}
