import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import { resolveFile } from "@/lib/file-utils";

export async function POST(req: NextRequest) {
  try {
    const { kosisPath, macroPath, step } = await req.json();
    // Resolve Blob URLs to local paths
    const localKosis = kosisPath ? await resolveFile(kosisPath) : "";
    const localMacro = macroPath ? await resolveFile(macroPath) : "";

    // Step-based execution
    if (step === 1) {
      // Step 1: Validate KOSIS file
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(localKosis);
      const sheetNames = wb.worksheets.map(s => s.name);
      const rowCount = wb.worksheets[0]?.rowCount || 0;
      return NextResponse.json({
        success: true,
        step: 1,
        output: `KOSIS 파일 로드 완료\n시트: ${sheetNames.join(", ")}\n행 수: ${rowCount}`
      });
    }

    if (step === 2) {
      // Step 2: Read KOSIS data
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(localKosis);
      const ws = wb.worksheets[0];
      if (!ws) throw new Error("시트가 없습니다");

      // Read headers (row 1) and find the latest month columns
      const headers: string[] = [];
      ws.getRow(1).eachCell((cell, col) => { headers[col] = String(cell.value || ""); });

      // Read data rows
      const data: Record<string, string[]> = {};
      ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const category = String(row.getCell(1).value || "");
        const values: string[] = [];
        row.eachCell((cell, col) => { if (col > 1) values.push(String(cell.value || "")); });
        data[category] = values;
      });

      return NextResponse.json({
        success: true,
        step: 2,
        output: `데이터 읽기 완료\n카테고리: ${Object.keys(data).length}개\n컬럼: ${headers.filter(h => h).length}개`,
        data: { headers, categories: Object.keys(data).slice(0, 10) }
      });
    }

    if (step === 3) {
      // Step 3: Open Macro Analysis
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(localMacro);
      const sheetNames = wb.worksheets.map(s => s.name);
      return NextResponse.json({
        success: true,
        step: 3,
        output: `Macro 엑셀 열기 완료\n시트 ${sheetNames.length}개: ${sheetNames.join(", ")}`
      });
    }

    if (step === 4) {
      // Step 4: Update all 10 sheets
      const kosisWb = new ExcelJS.Workbook();
      await kosisWb.xlsx.readFile(localKosis);
      const kosisWs = kosisWb.worksheets[0];

      const macroWb = new ExcelJS.Workbook();
      await macroWb.xlsx.readFile(localMacro);

      // Read KOSIS data into memory
      const kosisData: Record<string, (string | number | null)[]> = {};
      const kosisHeaders: string[] = [];
      if (kosisWs) {
        kosisWs.getRow(1).eachCell((cell, col) => { kosisHeaders[col] = String(cell.value || ""); });
        kosisWs.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const cat = String(row.getCell(1).value || "").trim();
          const vals: (string | number | null)[] = [];
          row.eachCell((cell, col) => { if (col > 1) vals[col] = cell.value as string | number | null; });
          kosisData[cat] = vals;
        });
      }

      let updatedSheets = 0;
      let updatedCells = 0;

      // For each Macro sheet, try to match and update
      for (const sheet of macroWb.worksheets) {
        const sheetName = sheet.name;
        // Find the last column with data in row 1 (header row)
        let lastCol = 0;
        sheet.getRow(1).eachCell((cell, col) => { if (cell.value) lastCol = col; });

        // The new data goes into lastCol + 1
        const newCol = lastCol + 1;

        // Match rows by category name (column 1 or 2)
        let matched = false;
        sheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return; // skip header
          const cat = String(row.getCell(1).value || row.getCell(2).value || "").trim();
          if (kosisData[cat]) {
            // Find the latest value from KOSIS
            const vals = kosisData[cat];
            const lastVal = vals[vals.length - 1] ?? vals.filter(v => v != null).pop();
            if (lastVal !== undefined) {
              row.getCell(newCol).value = typeof lastVal === 'number' ? lastVal : Number(lastVal) || lastVal;
              updatedCells++;
              matched = true;
            }
          }
        });
        if (matched) updatedSheets++;
      }

      // Save
      const outputPath = `/tmp/${Date.now()}_macro_updated.xlsx`;
      await macroWb.xlsx.writeFile(outputPath);

      return NextResponse.json({
        success: true,
        step: 4,
        output: `업데이트 완료\n수정된 시트: ${updatedSheets}개\n수정된 셀: ${updatedCells}개`,
        outputPath
      });
    }

    if (step === 5) {
      // Step 5: Confirm save (just return the output path)
      return NextResponse.json({
        success: true,
        step: 5,
        output: `저장 완료. 다운로드 버튼을 눌러 결과 파일을 받으세요.`
      });
    }

    return NextResponse.json({ error: "Invalid step" }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
