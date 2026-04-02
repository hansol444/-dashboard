import { NextRequest, NextResponse } from "next/server";
import { exec } from "child_process";
import { promisify } from "util";
import { writeFile, mkdir } from "fs/promises";
import path from "path";

const execAsync = promisify(exec);
const WORKSPACE = path.resolve(process.cwd(), "..");
const MEETING_INPUT = path.join(WORKSPACE, "meeting-notes", "input");

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    if (!file || !file.name.endsWith(".txt")) {
      return NextResponse.json({ error: "TXT 파일을 업로드하세요." }, { status: 400 });
    }

    // 1. TXT 파일을 meeting-notes/input/에 저장
    await mkdir(MEETING_INPUT, { recursive: true });
    const buffer = Buffer.from(await file.arrayBuffer());
    const filePath = path.join(MEETING_INPUT, file.name);
    await writeFile(filePath, buffer);

    // 2. summarize.py --notion 실행
    const command = `python meeting-notes/summarize.py "meeting-notes/input/${file.name}" --notion`;
    const { stdout, stderr } = await execAsync(command, {
      cwd: WORKSPACE,
      timeout: 300_000,
      env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONUTF8: "1" },
      encoding: "utf-8",
    });

    // 3. stdout에서 __RESULT_JSON__ 마커로 구조화 결과 파싱
    const jsonMatch = stdout.match(/__RESULT_JSON__(.+?)__END_JSON__/);
    if (!jsonMatch) {
      return NextResponse.json({
        success: true,
        parsed: false,
        stdout: stdout.slice(0, 2000),
        stderr: stderr.slice(0, 500),
      });
    }

    const result = JSON.parse(jsonMatch[1]);

    return NextResponse.json({
      success: true,
      parsed: true,
      summary: result.summary || "",
      participants: result.participants || [],
      meetingDate: result.meeting_date || "",
      taskAssignments: (result.task_assignments || []).map((ta: { assignee: string; task: string; deadline?: string }) => ({
        assignee: ta.assignee,
        task: ta.task,
        deadline: ta.deadline || "미정",
      })),
      directionChanges: (result.direction_changes || []).map((dc: { from_who: string; content: string }) => ({
        fromWho: dc.from_who,
        content: dc.content,
      })),
      notionUrl: result.notion_url || "",
      actionItems: result.action_items_for_dashboard || [],
      stdout: stdout.slice(0, 500),
    });
  } catch (err: unknown) {
    const error = err as { stdout?: string; stderr?: string; message?: string };
    return NextResponse.json({
      success: false,
      error: error.message ?? "실행 실패",
      stdout: error.stdout?.slice(0, 1000) ?? "",
      stderr: error.stderr?.slice(0, 500) ?? "",
    }, { status: 500 });
  }
}
