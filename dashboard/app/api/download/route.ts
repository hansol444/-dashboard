import { NextRequest, NextResponse } from "next/server";
import { readFile } from "fs/promises";
import path from "path";

export async function GET(req: NextRequest) {
  const filePath = req.nextUrl.searchParams.get("path");
  if (!filePath || !filePath.startsWith("/tmp/")) {
    return NextResponse.json({ error: "Invalid path" }, { status: 400 });
  }

  try {
    const buffer = await readFile(filePath);
    const filename = path.basename(filePath);
    const ext = path.extname(filename).toLowerCase();

    const contentTypes: Record<string, string> = {
      ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      ".csv": "text/csv",
      ".json": "application/json",
    };

    return new NextResponse(buffer, {
      headers: {
        "Content-Type": contentTypes[ext] || "application/octet-stream",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(filename)}"`,
      },
    });
  } catch {
    return NextResponse.json({ error: "File not found" }, { status: 404 });
  }
}
