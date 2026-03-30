import { NextRequest, NextResponse } from "next/server";
import { getAllTasks, upsertTask, removeTask, patchTask } from "../../../lib/github-db";

/** GET: 전체 태스크 반환 (대시보드 폴링) */
export async function GET() {
  try {
    const tasks = await getAllTasks();
    return NextResponse.json({ tasks });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST: 태스크 추가/업데이트 (수동 입력 또는 내부 호출) */
export async function POST(request: NextRequest) {
  try {
    const task = await request.json();
    await upsertTask(task);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH: 상태/마감기한 등 부분 업데이트 */
export async function PATCH(request: NextRequest) {
  try {
    const { id, ...patch } = await request.json();
    await patchTask(id, patch);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE: 태스크 삭제 */
export async function DELETE(request: NextRequest) {
  try {
    const { id } = await request.json();
    await removeTask(id);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
