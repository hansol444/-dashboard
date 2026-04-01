import { NextRequest, NextResponse } from "next/server";
import { writeFile } from "fs/promises";
import path from "path";

export async function POST(req: NextRequest) {
  const formData = await req.formData();
  const file = formData.get("file") as File | null;
  if (!file) return NextResponse.json({ error: "No file" }, { status: 400 });

  const bytes = await file.arrayBuffer();
  const buffer = Buffer.from(bytes);
  const tmpPath = path.join("/tmp", `${Date.now()}_${file.name}`);
  await writeFile(tmpPath, buffer);

  return NextResponse.json({ success: true, path: tmpPath, filename: file.name });
}
