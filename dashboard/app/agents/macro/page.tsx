"use client";

import { useState, useRef } from "react";
import { upload } from "@vercel/blob/client";
import Link from "next/link";

interface Step {
  label: string;
  description: string;
  status: "locked" | "ready" | "running" | "done" | "error";
  output?: string;
}

export default function MacroAgent() {
  const [kosisFile, setKosisFile] = useState<File | null>(null);
  const [macroFile, setMacroFile] = useState<File | null>(null);
  const [kosisPath, setKosisPath] = useState("");
  const [macroPath, setMacroPath] = useState("");
  const [outputPath, setOutputPath] = useState("");
  const [filesUploaded, setFilesUploaded] = useState(false);
  const [uploading, setUploading] = useState(false);
  const kosisRef = useRef<HTMLInputElement>(null);
  const macroRef = useRef<HTMLInputElement>(null);

  const [steps, setSteps] = useState<Step[]>([
    { label: "KOSIS 파일 검증", description: "업로드된 KOSIS 파일의 시트 구조와 데이터를 확인합니다.", status: "locked" },
    { label: "데이터 읽기", description: "KOSIS 파일에서 카테고리별 지표 데이터를 읽습니다.", status: "locked" },
    { label: "Macro 엑셀 열기", description: "Macro Analysis 엑셀 워크북의 시트 목록을 확인합니다.", status: "locked" },
    { label: "10개 시트 업데이트", description: "빈일자리·채용·근로자·입직자 (상용/임시일용) 10개 시트에 데이터를 매칭하여 입력합니다.", status: "locked" },
    { label: "저장 및 다운로드", description: "업데이트된 엑셀 파일을 저장하고 다운로드합니다.", status: "locked" },
  ]);
  const [running, setRunning] = useState(false);

  const uploadFiles = async () => {
    if (!kosisFile || !macroFile) return;
    setUploading(true);

    try {
      const uploadOne = async (file: File) => {
        const blob = await upload(file.name, file, {
          access: "public",
          handleUploadUrl: "/api/upload",
        });
        return blob;
      };

      const [kosisBlob, macroBlob] = await Promise.all([uploadOne(kosisFile), uploadOne(macroFile)]);

      setKosisPath(kosisBlob.url);
      setMacroPath(macroBlob.url);
      setFilesUploaded(true);
      setSteps((prev) => prev.map((s, i) => (i === 0 ? { ...s, status: "ready" } : s)));
    } catch (err) {
      console.error(err);
    }
    setUploading(false);
  };

  const runStep = async (stepIdx: number) => {
    setRunning(true);
    setSteps((prev) => prev.map((s, i) => (i === stepIdx ? { ...s, status: "running", output: undefined } : s)));

    try {
      const res = await fetch("/api/agents/macro", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kosisPath, macroPath, step: stepIdx + 1 }),
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
          <h1 className="text-2xl font-bold">📊 Macro 분석</h1>
          <span className="text-xs px-2 py-0.5 rounded-full bg-green-500/20 text-green-400">● 활성</span>
        </div>
        <p className="text-sm" style={{ color: "var(--text-muted)" }}>
          KOSIS 데이터를 Macro Analysis 엑셀 10개 시트에 자동 반영합니다.
        </p>
      </div>

      {/* 파일 업로드 */}
      <div className={`p-5 rounded-xl mb-6 border ${filesUploaded ? "border-green-500/30 bg-green-500/5" : "border-gray-600"}`} style={filesUploaded ? {} : { background: "var(--surface)" }}>
        <h2 className="text-sm font-medium mb-3" style={{ color: "var(--text-muted)" }}>파일 업로드</h2>

        <div className="space-y-3">
          <div>
            <label className="text-xs text-gray-400 block mb-1">1. KOSIS 파일 (산업_규모별_고용_*.xlsx)</label>
            <div className="flex gap-3">
              <input ref={kosisRef} type="file" accept=".xlsx,.xls" className="hidden"
                onChange={(e) => setKosisFile(e.target.files?.[0] || null)} />
              <button onClick={() => kosisRef.current?.click()} disabled={filesUploaded}
                className="flex-1 px-4 py-2.5 rounded-lg text-sm text-left bg-black border border-gray-700 text-white disabled:opacity-50 hover:border-gray-500 transition-all">
                {kosisFile ? `📄 ${kosisFile.name}` : "파일 선택..."}
              </button>
            </div>
          </div>

          <div>
            <label className="text-xs text-gray-400 block mb-1">2. Macro Analysis 엑셀</label>
            <div className="flex gap-3">
              <input ref={macroRef} type="file" accept=".xlsx,.xls" className="hidden"
                onChange={(e) => setMacroFile(e.target.files?.[0] || null)} />
              <button onClick={() => macroRef.current?.click()} disabled={filesUploaded}
                className="flex-1 px-4 py-2.5 rounded-lg text-sm text-left bg-black border border-gray-700 text-white disabled:opacity-50 hover:border-gray-500 transition-all">
                {macroFile ? `📄 ${macroFile.name}` : "파일 선택..."}
              </button>
            </div>
          </div>
        </div>

        {!filesUploaded ? (
          <button onClick={uploadFiles} disabled={!kosisFile || !macroFile || uploading}
            className="mt-4 w-full px-6 py-2.5 rounded-lg text-sm font-medium text-black hover:opacity-80 disabled:opacity-30"
            style={{ background: "var(--accent)" }}>
            {uploading ? "⏳ 업로드 중..." : "업로드"}
          </button>
        ) : (
          <div className="mt-3 text-sm text-green-400">✅ 파일 업로드 완료</div>
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
          <div className="text-green-400 font-medium mb-1">✅ Macro 업데이트 전체 완료</div>
          {outputPath && (
            <a href={`/api/download?path=${encodeURIComponent(outputPath)}`}
              className="inline-block mt-3 px-6 py-2.5 rounded-lg text-sm font-medium text-black hover:opacity-80"
              style={{ background: "var(--accent)" }}>
              📥 결과 엑셀 다운로드
            </a>
          )}
        </div>
      )}
    </div>
  );
}
