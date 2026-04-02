import { NextRequest, NextResponse } from "next/server";
import { put } from "@vercel/blob";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    if (!file) return NextResponse.json({ error: "No file" }, { status: 400 });

    // Vercel Blob에 업로드 (크기 제한 없음)
    const blob = await put(`agents/${Date.now()}_${file.name}`, file, {
      access: "public",
    });

    return NextResponse.json({
      success: true,
      path: blob.url,         // Blob URL
      filename: file.name,
      size: file.size,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
