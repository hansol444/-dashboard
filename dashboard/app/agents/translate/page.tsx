"use client";

import { useState, useRef } from "react";
import Link from "next/link";

interface Step {
  label: string;
  description: string;
  status: "locked" | "ready" | "running" | "done" | "error";
  output?: string;
}

export default function TranslateAgent() {
  const [pptxFile, setPptxFile] = useState<File | null>(null);
  const [pptxPath, setPptxPath] = useState("");
  const [outputPath, setOutputPath] = useState("");
  const [fileUploaded, setFileUploaded] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);

  const [steps, setSteps] = useState<Step[]>([
    { label: "PPT 파일 로드", description: "업로드된 PPTX 파일 정보를 확인합니다.", status: "locked" },
    { label: "텍스트 추출", description: "슬라이드별 텍스트 블록을 추출합니다.", status: "locked" },
    { label: "용어집 로드", description: "HR/Recruitment 도메인 용어집 + 보존어를 로드합니다.", status: "locked" },
    { label: "Claude API 번역", description: "슬라이드별 한→영 번역을 실행합니다 (Australian English).", status: "locked" },
    { label: "후처리 적용", description: "호주 영어 철자, 억/조 단위, 통화 순서 등 7개 규칙 적용.", status: "locked" },
    { label: "번역 PPT 저장", description: "번역 결과를 PPTX로 생성합니다.", status: "locked" },
  ]);
  const [running, setRunning] = useState(false);

  const uploadFile = async () => {
    if (!pptxFile) return;
    setUploading(true);
    setUploadError("");
    try {
      const form = new FormData();
      form.append("file", pptxFile);
      const res = await fetch("/api/upload", { method: "POST", body: form });
      if (!res.ok) {
        const text = await res.text();
        setUploadError(`업로드 실패 (HTTP ${res.status}): ${text.slice(0, 200)}`);
        setUploading(false);
        return;
      }
      const data = await res.json();
      if (data.success) {
        setPptxPath(data.path);
        setFileUploaded(true);
        setSteps((prev) => prev.map((s, i) => (i === 0 ? { ...s, status: "ready" } : s)));
      } else {
        setUploadError(data.error || `업로드 실패 (HTTP ${res.status})`);
      }
    } catch (err) {
      setUploadError(`업로드 에러: ${err}`);
    }
    setUploading(false);
  };

  const runStep = async (stepIdx: number) => {
    setRunning(true);
    setSteps((prev) => prev.map((s, i) => (i === stepIdx ? { ...s, status: "running", output: undefined } : s)));

    try {
      const res = await fetch("/api/agents/translate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pptxPath, step: stepIdx + 1 }),
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
          <h1 className="text-2xl font-bold">🌐 장표 번역</h1>
          <span className="text-xs px-2 py-0.5 rounded-full bg-green-500/20 text-green-400">● 활성</span>
        </div>
        <p className="text-sm" style={{ color: "var(--text-muted)" }}>
          한글 PPT → 영문 PPT 번역 (Australian English, 용어집 + 후처리)
        </p>
      </div>

      {/* 파일 업로드 */}
      <div className={`p-5 rounded-xl mb-6 border ${fileUploaded ? "border-green-500/30 bg-green-500/5" : "border-gray-600"}`} style={fileUploaded ? {} : { background: "var(--surface)" }}>
        <h2 className="text-sm font-medium mb-3" style={{ color: "var(--text-muted)" }}>번역할 PPT 업로드</h2>
        <div className="flex gap-3">
          <input ref={fileRef} type="file" accept=".pptx" className="hidden"
            onChange={(e) => setPptxFile(e.target.files?.[0] || null)} />
          <button onClick={() => fileRef.current?.click()} disabled={fileUploaded}
            className="flex-1 px-4 py-2.5 rounded-lg text-sm text-left bg-black border border-gray-700 text-white disabled:opacity-50 hover:border-gray-500 transition-all">
            {pptxFile ? `📄 ${pptxFile.name}` : "PPTX 파일 선택..."}
          </button>
          {!fileUploaded ? (
            <button onClick={uploadFile} disabled={!pptxFile || uploading}
              className="px-6 py-2.5 rounded-lg text-sm font-medium text-black hover:opacity-80 disabled:opacity-30"
              style={{ background: "var(--accent)" }}>
              {uploading ? "⏳ 업로드 중..." : "업로드"}
            </button>
          ) : (
            <span className="px-6 py-2.5 text-green-400 text-sm">✅</span>
          )}
        </div>
        {uploadError && (
          <div className="mt-3 p-3 rounded-lg text-xs font-mono bg-red-500/10 text-red-300">
            {uploadError}
          </div>
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
          <div className="text-green-400 font-medium mb-1">✅ 번역 완료</div>
          {outputPath && (
            <a href={`/api/download?path=${encodeURIComponent(outputPath)}`}
              className="inline-block mt-3 px-6 py-2.5 rounded-lg text-sm font-medium text-black hover:opacity-80"
              style={{ background: "var(--accent)" }}>
              📥 번역 PPT 다운로드
            </a>
          )}
        </div>
      )}
    </div>
  );
}
