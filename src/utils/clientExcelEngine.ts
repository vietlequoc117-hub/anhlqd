import ExcelJS from "exceljs";
import {
  ExcelSheetData,
  ExcelHeaderCol,
  ExcelStudentRow,
  normalizeSbdClient,
  canonicalSbdKey,
  parseVietnameseScoreClient,
} from "./sampleGradingSheet";

export function colNumberToLetter(col: number): string {
  let temp = 0;
  let letter = "";
  let c = col;
  while (c > 0) {
    temp = (c - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    c = (c - temp - 1) / 26;
  }
  return letter;
}

export function colLetterToNumber(letter: string): number {
  const clean = letter.trim().toUpperCase();
  let col = 0;
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col;
}

export function extractCellValue(cellValue: ExcelJS.CellValue): string | number | null {
  if (cellValue === null || cellValue === undefined) return null;
  if (typeof cellValue === "number" || typeof cellValue === "string") {
    return cellValue;
  }
  if (typeof cellValue === "boolean") {
    return cellValue ? "TRUE" : "FALSE";
  }
  if (cellValue instanceof Date) {
    const d = String(cellValue.getDate()).padStart(2, "0");
    const m = String(cellValue.getMonth() + 1).padStart(2, "0");
    const y = cellValue.getFullYear();
    return `${d}/${m}/${y}`;
  }
  if (typeof cellValue === "object") {
    if ("richText" in cellValue && Array.isArray(cellValue.richText)) {
      return cellValue.richText.map((rt) => rt.text).join("");
    }
    if ("result" in cellValue && cellValue.result !== undefined) {
      if (typeof cellValue.result === "number" || typeof cellValue.result === "string") {
        return cellValue.result;
      }
    }
    if ("text" in cellValue && typeof cellValue.text === "string") {
      return cellValue.text;
    }
  }
  return String(cellValue);
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const clean = base64.includes(",") ? base64.split(",")[1] : base64;
  const binaryString = atob(clean);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(
      null,
      Array.from(bytes.subarray(i, i + chunkSize))
    );
  }
  return btoa(binary);
}

function detectSheetColumns(worksheet: ExcelJS.Worksheet): {
  headerRow: number;
  sbdCol: string;
  phachCol: string;
  scoreCol: string;
} {
  let detectedHeaderRow = 5;
  let detectedSbdCol = "G";
  let detectedPhachCol = "";
  let detectedScoreCol = "I";

  for (let r = 1; r <= Math.min(18, worksheet.rowCount); r++) {
    const row = worksheet.getRow(r);
    let foundSbd = "";
    let foundPhach = "";
    let foundScore = "";

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const text = String(extractCellValue(cell.value) ?? "")
        .replace(/[\r\n]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();

      if (
        text.includes("số bd") ||
        text.includes("sbd") ||
        text.includes("số báo danh") ||
        text === "mã sbd"
      ) {
        foundSbd = colNumberToLetter(colNumber);
      }
      if (
        text.includes("mã phách") ||
        text.includes("ma phach") ||
        text === "phách" ||
        text.includes("số phách")
      ) {
        foundPhach = colNumberToLetter(colNumber);
      }
      if (
        text === "điểm" ||
        text === "điểm số" ||
        text === "điểm thi" ||
        text.includes("điểm bằng số")
      ) {
        foundScore = colNumberToLetter(colNumber);
      }
    });

    if (foundSbd || foundPhach || foundScore) {
      detectedHeaderRow = r;
      if (foundSbd) detectedSbdCol = foundSbd;
      if (foundPhach) detectedPhachCol = foundPhach;
      if (foundScore) detectedScoreCol = foundScore;
      break;
    }
  }

  return {
    headerRow: detectedHeaderRow,
    sbdCol: detectedSbdCol,
    phachCol: detectedPhachCol,
    scoreCol: detectedScoreCol,
  };
}

/**
 * Automatically classifies whether a loaded Excel file is:
 * - "PHACH_MAP" (Hình 2: Bảng ghép SBD <-> MÃ PHÁCH, has a MÃ PHÁCH column)
 * - "ROSTER_SBD" (Hình 1: Bảng điểm thí sinh chính thức có Họ và tên, Ngày sinh, SBD, Điểm)
 */
export function detectExcelFileRole(
  data: ExcelSheetData
): "ROSTER_SBD" | "PHACH_MAP" {
  if (data.detectedPhachCol) {
    return "PHACH_MAP";
  }
  const hasPhachHeader = data.headers.some((h) => {
    const lbl = h.label.toLowerCase();
    return lbl.includes("phách") || lbl.includes("ma phach");
  });
  return hasPhachHeader ? "PHACH_MAP" : "ROSTER_SBD";
}

/**
 * 100% Client-Side Excel Inspector (Runs directly in browser on Vercel / GitHub Pages without backend API)
 * Supports both Hình 1 (Bảng điểm SBD) and Hình 2 (Bảng ghép SBD <-> MÃ PHÁCH).
 */
export async function parseExcelClientSide(params: {
  fileBase64: string;
  fileName: string;
  preferredHeaderRow?: number;
  preferredSbdCol?: string;
  preferredPhachCol?: string;
  preferredScoreCol?: string;
  targetSheetIndex?: number;
}): Promise<ExcelSheetData> {
  const {
    fileBase64,
    fileName,
    preferredHeaderRow,
    preferredSbdCol,
    preferredPhachCol,
    preferredScoreCol,
    targetSheetIndex,
  } = params;

  const arrayBuffer = base64ToArrayBuffer(fileBase64);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  const availableSheets = workbook.worksheets.map((ws, idx) => ({
    index: idx,
    name: ws.name,
    rowCount: ws.rowCount,
  }));

  const sheetIdx =
    typeof targetSheetIndex === "number" &&
    targetSheetIndex >= 0 &&
    targetSheetIndex < workbook.worksheets.length
      ? targetSheetIndex
      : 0;

  const worksheet = workbook.worksheets[sheetIdx] || workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("File Excel không có trang tính (worksheet) nào.");
  }

  const autoDetected = detectSheetColumns(worksheet);
  const detectedHeaderRow = preferredHeaderRow || autoDetected.headerRow;
  const detectedSbdCol = preferredSbdCol || autoDetected.sbdCol;
  const detectedPhachCol =
    preferredPhachCol !== undefined ? preferredPhachCol : autoDetected.phachCol;
  const detectedScoreCol = preferredScoreCol || autoDetected.scoreCol;

  const titleRows: Array<{ rowNumber: number; texts: string[] }> = [];
  for (let r = 1; r < detectedHeaderRow; r++) {
    const row = worksheet.getRow(r);
    const seenTexts = new Set<string>();
    row.eachCell({ includeEmpty: false }, (cell) => {
      const val = String(extractCellValue(cell.value) ?? "").trim();
      if (val && !seenTexts.has(val)) {
        seenTexts.add(val);
      }
    });
    if (seenTexts.size > 0) {
      titleRows.push({ rowNumber: r, texts: Array.from(seenTexts) });
    }
  }

  const headerRowObj = worksheet.getRow(detectedHeaderRow);
  const maxCol = Math.max(worksheet.columnCount, 6);
  const headers: ExcelHeaderCol[] = [];

  for (let c = 1; c <= maxCol; c++) {
    const colLetter = colNumberToLetter(c);
    const rawLabel = String(extractCellValue(headerRowObj.getCell(c).value) ?? "")
      .replace(/[\r\n]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const label = rawLabel === "ST T" ? "STT" : rawLabel;
    const colWidth = worksheet.getColumn(c).width || 12;
    if (label || c <= worksheet.columnCount) {
      headers.push({
        colLetter,
        colNumber: c,
        label: label || `Cột ${colLetter}`,
        width: Math.round(colWidth * 10) / 10,
      });
    }
  }

  const sbdColNum = colLetterToNumber(detectedSbdCol);
  const phachColNum = detectedPhachCol ? colLetterToNumber(detectedPhachCol) : 0;
  const scoreColNum = colLetterToNumber(detectedScoreCol);
  const studentRows: ExcelStudentRow[] = [];

  for (let r = detectedHeaderRow + 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const sbdRaw = extractCellValue(row.getCell(sbdColNum).value);
    const sbdStr = sbdRaw !== null ? String(sbdRaw).trim() : "";

    const phachRaw =
      phachColNum > 0 ? extractCellValue(row.getCell(phachColNum).value) : null;
    const phachStr = phachRaw !== null ? String(phachRaw).trim() : "";

    let nonEmptyCount = 0;
    const cells: Record<string, string | number | null> = {};
    for (const h of headers) {
      const cellVal = extractCellValue(row.getCell(h.colNumber).value);
      cells[h.colLetter] = cellVal;
      if (cellVal !== null && String(cellVal).trim() !== "") {
        nonEmptyCount++;
      }
    }

    if (nonEmptyCount === 0) continue;

    // Skip footer signature rows like "CHỮ KÝ CỦA CBCT" where SBD is empty or non-code
    const rowJoinedLower = Object.values(cells)
      .map((v) => String(v ?? "").toLowerCase())
      .join(" ");
    if (
      !sbdStr &&
      !phachStr &&
      (rowJoinedLower.includes("chữ ký") || rowJoinedLower.includes("cbct"))
    ) {
      continue;
    }

    const scoreCell = row.getCell(scoreColNum);
    const currentScore = extractCellValue(scoreCell.value);
    const fontName =
      scoreCell.font?.name || row.getCell(sbdColNum).font?.name || "Times New Roman";
    const hasBorder = Boolean(
      scoreCell.border &&
        (scoreCell.border.top ||
          scoreCell.border.bottom ||
          scoreCell.border.left ||
          scoreCell.border.right)
    );

    studentRows.push({
      rowNumber: r,
      sbd: sbdStr,
      normalizedSbd: normalizeSbdClient(sbdStr),
      maPhach: phachStr,
      normalizedPhach: normalizeSbdClient(phachStr),
      currentScore,
      cells,
      fontName,
      hasBorder,
    });
  }

  // Also scan all worksheets for SBD <-> MÃ PHÁCH pairs in case File Excel 2 has multiple sheets
  const allSheetsPhachPairs: Array<{
    sheetName: string;
    rowNumber: number;
    sbd: string;
    normalizedSbd: string;
    maPhach: string;
    normalizedPhach: string;
  }> = [];

  for (let wsIdx = 0; wsIdx < workbook.worksheets.length; wsIdx++) {
    const ws = workbook.worksheets[wsIdx];
    const cols =
      wsIdx === sheetIdx
        ? {
            headerRow: detectedHeaderRow,
            sbdCol: detectedSbdCol,
            phachCol: detectedPhachCol,
          }
        : detectSheetColumns(ws);

    if (!cols.sbdCol || !cols.phachCol) continue;
    const sCol = colLetterToNumber(cols.sbdCol);
    const pCol = colLetterToNumber(cols.phachCol);

    for (let r = cols.headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const rawSbd = extractCellValue(row.getCell(sCol).value);
      const rawPhach = extractCellValue(row.getCell(pCol).value);
      const sbdStr = rawSbd !== null ? String(rawSbd).trim() : "";
      const phachStr = rawPhach !== null ? String(rawPhach).trim() : "";
      if (!sbdStr && !phachStr) continue;
      const normSbd = normalizeSbdClient(sbdStr);
      if (!normSbd) continue;

      allSheetsPhachPairs.push({
        sheetName: ws.name,
        rowNumber: r,
        sbd: sbdStr,
        normalizedSbd: normSbd,
        maPhach: phachStr,
        normalizedPhach: normalizeSbdClient(phachStr),
      });
    }
  }

  const mergesCount =
    (worksheet.model && Array.isArray(worksheet.model.merges)
      ? worksheet.model.merges.length
      : 0) || 0;

  return {
    fileName: fileName || "Bang_Diem_Thi.xlsx",
    fileBase64,
    availableSheets,
    currentSheetIndex: sheetIdx,
    sheetName: worksheet.name,
    totalRows: worksheet.rowCount,
    mergedCellsCount: mergesCount,
    detectedHeaderRow,
    detectedSbdCol,
    detectedPhachCol: detectedPhachCol || undefined,
    detectedScoreCol,
    titleRows,
    headers,
    studentRows,
    allSheetsPhachPairs,
  };
}

/**
 * 100% Client-Side Non-Destructive Excel Updater (Preserves all formatting, borders, merges, fonts)
 * Supports matching directly by SBD OR via 3-way SBD <-> MÃ PHÁCH <-> ĐIỂM BẰNG SỐ lookup.
 */
export async function exportUpdatedExcelClientSide(params: {
  fileBase64: string;
  fileName: string;
  headerRow: number;
  sbdCol: string;
  scoreCol: string;
  sheetIndex?: number;
  ocrRecords: Array<{ sbd: string; maPhach?: string; score: number }>;
  sbdToPhachMap?: Record<string, string>;
  phachToSbdMap?: Record<string, string>;
}) {
  const {
    fileBase64,
    fileName,
    headerRow,
    sbdCol,
    scoreCol,
    sheetIndex,
    ocrRecords,
    sbdToPhachMap,
    phachToSbdMap,
  } = params;

  const arrayBuffer = base64ToArrayBuffer(fileBase64);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  const targetIdx =
    typeof sheetIndex === "number" &&
    sheetIndex >= 0 &&
    sheetIndex < workbook.worksheets.length
      ? sheetIndex
      : 0;
  const worksheet = workbook.worksheets[targetIdx] || workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("File Excel không hợp lệ.");
  }

  // Build score lookup map indexed by SBD, canonical SBD, Mã Phách, and canonical Mã Phách
  const ocrMap: Record<string, number> = {};
  for (const rec of ocrRecords || []) {
    const val = parseVietnameseScoreClient(rec.score);
    if (val === null) continue;

    const rawCode = rec.sbd;
    const normCode = normalizeSbdClient(rawCode);
    const canonCode = canonicalSbdKey(rawCode);

    if (normCode) ocrMap[normCode] = val;
    if (canonCode) ocrMap[canonCode] = val;

    if (rec.maPhach) {
      const normP = normalizeSbdClient(rec.maPhach);
      const canonP = canonicalSbdKey(rec.maPhach);
      if (normP) ocrMap[normP] = val;
      if (canonP) ocrMap[canonP] = val;
      if (phachToSbdMap) {
        const mappedSbd = phachToSbdMap[normP] || phachToSbdMap[canonP];
        if (mappedSbd) {
          ocrMap[normalizeSbdClient(mappedSbd)] = val;
          ocrMap[canonicalSbdKey(mappedSbd)] = val;
        }
      }
    }

    // Also check if `rec.sbd` itself is actually a Mã Phách code in `phachToSbdMap`
    if (phachToSbdMap) {
      const mappedSbd = phachToSbdMap[normCode] || phachToSbdMap[canonCode];
      if (mappedSbd) {
        ocrMap[normalizeSbdClient(mappedSbd)] = val;
        ocrMap[canonicalSbdKey(mappedSbd)] = val;
      }
    }
  }

  const sbdColNum = colLetterToNumber(sbdCol || "G");
  const scoreColNum = colLetterToNumber(scoreCol || "I");
  const startRow = (headerRow || 5) + 1;

  let matchedCount = 0;
  let untouchedCount = 0;
  const matchedDetails: Array<{
    rowNumber: number;
    sbd: string;
    maPhach?: string;
    score: number;
  }> = [];

  for (let r = startRow; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const rawSbd = extractCellValue(row.getCell(sbdColNum).value);
    const normSbd = normalizeSbdClient(rawSbd);
    const canonSbd = canonicalSbdKey(rawSbd);
    if (!normSbd && !canonSbd) continue;

    const linkedPhach =
      sbdToPhachMap?.[normSbd] || sbdToPhachMap?.[canonSbd] || "";
    const normPhach = normalizeSbdClient(linkedPhach);
    const canonPhach = canonicalSbdKey(linkedPhach);

    // If sbdToPhachMap is active and this student exists in the roster with an explicitly blank Mã Phách (like K9.002 in Hình 2),
    // only match if they have a valid SBD direct score or a non-empty Mã Phách score
    const matchedVal =
      (normSbd && Object.prototype.hasOwnProperty.call(ocrMap, normSbd)
        ? ocrMap[normSbd]
        : undefined) ??
      (canonSbd && Object.prototype.hasOwnProperty.call(ocrMap, canonSbd)
        ? ocrMap[canonSbd]
        : undefined) ??
      (normPhach && Object.prototype.hasOwnProperty.call(ocrMap, normPhach)
        ? ocrMap[normPhach]
        : undefined) ??
      (canonPhach && Object.prototype.hasOwnProperty.call(ocrMap, canonPhach)
        ? ocrMap[canonPhach]
        : undefined);

    if (matchedVal !== undefined) {
      const scoreCell = row.getCell(scoreColNum);
      const numericScore = Number(matchedVal);
      // Strictly only mutate .value as Float/Number, preserving all cell styles/borders/fonts
      scoreCell.value = numericScore;
      matchedCount++;
      matchedDetails.push({
        rowNumber: r,
        sbd: String(rawSbd).trim(),
        maPhach: linkedPhach || undefined,
        score: numericScore,
      });
    } else {
      untouchedCount++;
    }
  }

  const outBuffer = await workbook.xlsx.writeBuffer();
  const updatedExcelBase64 = arrayBufferToBase64(outBuffer);

  const mergesCount =
    (worksheet.model && Array.isArray(worksheet.model.merges)
      ? worksheet.model.merges.length
      : 0) || 0;

  const auditLines = [
    `[Kiểm định Cấu trúc Bảng tính Sau Khi So Khớp SBD ↔ MÃ PHÁCH ↔ ĐIỂM]`,
    ` - Trang tính đã xử lý: ${worksheet.name}`,
    ` - Số vùng gộp ô (Merged Cells) được bảo toàn: ${mergesCount}`,
    ` - Số ô cột ${scoreCol} (Điểm) đã ghi kiểu Float/Number: ${matchedCount}`,
    ...matchedDetails
      .slice(0, 12)
      .map(
        (d) =>
          `   + Hàng ${d.rowNumber} | SBD ${d.sbd}${
            d.maPhach ? ` (Mã phách: ${d.maPhach})` : ""
          } -> Ô ${scoreCol}${d.rowNumber} = ${d.score}`
      ),
  ];

  const entries = matchedDetails
    .map(
      (d) =>
        `    "${normalizeSbdClient(d.sbd)}": ${d.score},${
          d.maPhach ? `  # Mã phách: ${d.maPhach}` : ""
        }`
    )
    .join("\n");

  const openpyxlCode = `import re
import openpyxl

def normalize_code(value) -> str:
    """Chuẩn hóa mã Số Báo Danh / Mã Phách."""
    if value is None:
        return ""
    return re.sub(r"\\s+", "", str(value).strip()).replace(",", ".").upper()

def update_exam_scores(input_excel_path: str, output_excel_path: str):
    # 1. Dữ liệu (SBD -> Điểm số Float) sau khi khớp SBD (Hình 1) <-> MÃ PHÁCH (Hình 2) <-> ĐIỂM BẰNG SỐ (Hình 3)
    matched_sbd_scores: dict[str, float] = {
${entries || "    # Chưa có dữ liệu điểm"}
    }

    # 2. Mở file Excel nguồn (Hình 1) — giữ nguyên 100% định dạng, merge cells, borders, font
    wb = openpyxl.load_workbook(input_excel_path, data_only=False)
    sheet_names = wb.sheetnames
    ws = wb[sheet_names[${targetIdx}]] if len(sheet_names) > ${targetIdx} else wb.active

    START_ROW = ${startRow}
    SBD_COL = "${sbdCol}"
    SCORE_COL = "${scoreCol}"

    matched_count = 0

    # 3. Quét cột SBD (${sbdCol}) từ hàng ${startRow} trở đi và ghi điểm vào cột Điểm (${scoreCol}) thành Hình 4
    for row_idx in range(START_ROW, ws.max_row + 1):
        raw_sbd = ws[f"{SBD_COL}{row_idx}"].value
        norm_sbd = normalize_code(raw_sbd)

        if norm_sbd and norm_sbd in matched_sbd_scores:
            ws[f"{SCORE_COL}{row_idx}"].value = float(matched_sbd_scores[norm_sbd])
            matched_count += 1

    wb.save(output_excel_path)
    print(f"Hoàn tất! Đã khớp SBD - Mã Phách và điền điểm {matched_count} học sinh vào cột {SCORE_COL}.")

if __name__ == "__main__":
    update_exam_scores("${fileName}", "Da_Nhap_Diem_${fileName}")
`;

  return {
    updatedExcelBase64,
    outputFileName: `Da_Nhap_Diem_${fileName || "Bang_Diem_Thi_THCS_My_Xa.xlsx"}`,
    matchedCount,
    untouchedCount,
    matchedDetails,
    pythonAuditLog: auditLines.join("\n"),
    openpyxlCode,
  };
}

/**
 * Generates the exact 2-Excel Sample Bundle from the user's screenshots:
 * - Hình 1: BẢNG ĐIỂM THI MÔN: ANH 9 (K9.001 .. K9.005, Cột G = SBD, Cột I = Điểm)
 * - Hình 2: PHIẾU CHẤM BÀI THI MÔN TIẾNG ANH 9 (Cột B = SBD, Cột C = MÃ PHÁCH: A40, trống, A41, A42, A43)
 */
export async function generateSampleAnh9PhachBundleClientSide(): Promise<{
  rosterExcel: ExcelSheetData;
  phachExcel: ExcelSheetData;
}> {
  // 1. Build Hình 1: BẢNG ĐIỂM THI MÔN: ANH 9
  const wb1 = new ExcelJS.Workbook();
  const ws1 = wb1.addWorksheet("Bang_Diem_Anh_9", {
    views: [{ showGridLines: true }],
  });

  ws1.columns = [
    { key: "stt", width: 6 },        // A: STT
    { key: "hoten", width: 24 },     // B: Họ và tên
    { key: "ngaysinh", width: 15 },  // C: Ngày sinh
    { key: "gioi", width: 9 },       // D: Giới
    { key: "phongthi", width: 11 },  // E: Phòng thi
    { key: "diadiem", width: 12 },   // F: Địa điểm thi
    { key: "sbd", width: 14 },       // G: SBD
    { key: "made", width: 10 },      // H: Mã đề
    { key: "diem", width: 14 },      // I: Điểm
    { key: "ghichu", width: 12 },    // J: Ghi chú
  ];

  ws1.mergeCells("A1:D1");
  ws1.getCell("A1").value = "SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH NINH BÌNH";
  ws1.getCell("A1").font = { name: "Times New Roman", size: 12, bold: false };
  ws1.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };

  ws1.mergeCells("A2:D2");
  ws1.getCell("A2").value = "TRƯỜNG THCS MỸ XÁ";
  ws1.getCell("A2").font = { name: "Times New Roman", size: 12, bold: true };
  ws1.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  ws1.mergeCells("A3:J3");
  ws1.getCell("A3").value = "BẢNG ĐIỂM THI MÔN: ANH 9";
  ws1.getCell("A3").font = { name: "Times New Roman", size: 13, bold: true };
  ws1.getCell("A3").alignment = { horizontal: "center", vertical: "middle" };
  ws1.getRow(3).height = 24;

  const headerTitles1 = [
    "STT",
    "Họ và tên",
    "Ngày sinh",
    "Giới",
    "Phòng thi",
    "Địa điểm thi",
    "SBD",
    "Mã đề",
    "Điểm",
    "Ghi chú",
  ];
  const headerRow1 = ws1.getRow(6);
  headerRow1.height = 32;
  headerTitles1.forEach((title, idx) => {
    const cell = headerRow1.getCell(idx + 1);
    cell.value = title;
    cell.font = { name: "Times New Roman", size: 11, bold: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFEDEDED" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
  });

  const students1 = [
    { stt: 1, hoten: "Đinh Thiện An", ngaysinh: "20/10/2012", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K9.001", made: "", ghichu: "9A1" },
    { stt: 2, hoten: "Phùng Hà An", ngaysinh: "22/05/2012", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K9.002", made: "", ghichu: "9A5" },
    { stt: 3, hoten: "Trần Khánh An", ngaysinh: "10/01/2012", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K9.003", made: "", ghichu: "9A5" },
    { stt: 4, hoten: "Trần Ngọc Bảo An", ngaysinh: "02/07/2012", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K9.004", made: "", ghichu: "9A4" },
    { stt: 5, hoten: "Trịnh Đặng Bảo An", ngaysinh: "25/02/2012", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K9.005", made: "", ghichu: "9A6" },
  ];

  students1.forEach((student, index) => {
    const rowNum = 7 + index;
    const row = ws1.getRow(rowNum);
    row.height = 24;
    const values = [
      student.stt,
      student.hoten,
      student.ngaysinh,
      student.gioi,
      student.phong,
      student.diadiem,
      student.sbd,
      student.made,
      null, // Cột I: Điểm (Để trống như Hình 1, sẽ điền thành Hình 4)
      student.ghichu,
    ];
    values.forEach((val, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = val;
      cell.font = { name: "Times New Roman", size: 11, bold: false };
      cell.alignment = {
        horizontal: colIdx === 1 ? "left" : "center",
        vertical: "middle",
      };
      if (colIdx === 8) {
        cell.numFmt = "0.#";
      }
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
  });

  // 2. Build Hình 2: PHIẾU CHẤM BÀI THI - MÔN TIẾNG ANH 9 (Có cột SBD & MÃ PHÁCH)
  const wb2 = new ExcelJS.Workbook();
  const ws2 = wb2.addWorksheet("Phieu_Cham_Ma_Phach", {
    views: [{ showGridLines: true }],
  });

  ws2.columns = [
    { key: "stt", width: 8 },         // A: STT
    { key: "sbd", width: 16 },        // B: SBD
    { key: "maphach", width: 16 },    // C: MÃ PHÁCH
    { key: "diemso", width: 20 },     // D: ĐIỂM BẰNG SỐ
    { key: "diemchu", width: 24 },    // E: ĐIỂM BẰNG CHỮ
    { key: "ghichu", width: 18 },     // F: GHI CHÚ
  ];

  ws2.mergeCells("A1:D1");
  ws2.getCell("A1").value = "SỞ GD&ĐT NINH BÌNH";
  ws2.getCell("A1").font = { name: "Times New Roman", size: 12, bold: false };
  ws2.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };

  ws2.getCell("E1").value = "PHIẾU CHẤM BÀI THI";
  ws2.getCell("E1").font = { name: "Times New Roman", size: 12, bold: true };
  ws2.getCell("E1").alignment = { horizontal: "center", vertical: "middle" };

  ws2.mergeCells("A2:D2");
  ws2.getCell("A2").value = "TRƯỜNG THCS MỸ XÁ";
  ws2.getCell("A2").font = { name: "Times New Roman", size: 12, bold: true, underline: true };
  ws2.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  ws2.getCell("E2").value = "MÔN TIẾNG ANH 9";
  ws2.getCell("E2").font = { name: "Times New Roman", size: 12, bold: true, underline: true };
  ws2.getCell("E2").alignment = { horizontal: "center", vertical: "middle" };

  const headerTitles2 = [
    "STT",
    "SBD",
    "MÃ PHÁCH",
    "ĐIỂM BẰNG SỐ",
    "ĐIỂM BẰNG CHỮ",
    "GHI CHÚ",
  ];
  const headerRow2 = ws2.getRow(4);
  headerRow2.height = 28;
  headerTitles2.forEach((title, idx) => {
    const cell = headerRow2.getCell(idx + 1);
    cell.value = title;
    cell.font = { name: "Times New Roman", size: 11, bold: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFEDEDED" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
  });

  const phachRows = [
    { stt: 1, sbd: "K9.001", maPhach: "A40" },
    { stt: 2, sbd: "K9.002", maPhach: "" },
    { stt: 3, sbd: "K9.003", maPhach: "A41" },
    { stt: 4, sbd: "K9.004", maPhach: "A42" },
    { stt: 5, sbd: "K9.005", maPhach: "A43" },
  ];

  phachRows.forEach((item, idx) => {
    const row = ws2.getRow(5 + idx);
    row.height = 28;
    const vals = [item.stt, item.sbd, item.maPhach || null, null, null, null];
    vals.forEach((v, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = v;
      cell.font = { name: "Times New Roman", size: 14, bold: false };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });
  });

  const buf1 = await wb1.xlsx.writeBuffer();
  const buf2 = await wb2.xlsx.writeBuffer();

  const rosterExcel = await parseExcelClientSide({
    fileBase64: arrayBufferToBase64(buf1),
    fileName: "Hinh_1_Bang_Diem_Thi_Anh_9.xlsx",
    preferredHeaderRow: 6,
    preferredSbdCol: "G",
    preferredScoreCol: "I",
  });

  const phachExcel = await parseExcelClientSide({
    fileBase64: arrayBufferToBase64(buf2),
    fileName: "Hinh_2_Bang_Ma_Phach_Anh_9.xlsx",
    preferredHeaderRow: 4,
    preferredSbdCol: "B",
    preferredPhachCol: "C",
    preferredScoreCol: "D",
  });

  return { rosterExcel, phachExcel };
}

/**
 * Generates the single-Excel sample workbook (Môn Văn 6)
 */
export async function generateSampleExcelClientSide(): Promise<ExcelSheetData> {
  const { rosterExcel } = await generateSampleAnh9PhachBundleClientSide();
  return rosterExcel;
}
