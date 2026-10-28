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

/**
 * 100% Client-Side Excel Inspector (Runs directly in browser on Vercel / GitHub Pages without backend API)
 */
export async function parseExcelClientSide(params: {
  fileBase64: string;
  fileName: string;
  preferredHeaderRow?: number;
  preferredSbdCol?: string;
  preferredScoreCol?: string;
  targetSheetIndex?: number;
}): Promise<ExcelSheetData> {
  const {
    fileBase64,
    fileName,
    preferredHeaderRow,
    preferredSbdCol,
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

  let detectedHeaderRow = preferredHeaderRow || 5;
  let detectedSbdCol = preferredSbdCol || "G";
  let detectedScoreCol = preferredScoreCol || "I";

  if (!preferredHeaderRow) {
    for (let r = 1; r <= Math.min(15, worksheet.rowCount); r++) {
      const row = worksheet.getRow(r);
      let foundSbd = "";
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
          text === "điểm" ||
          text === "điểm số" ||
          text === "điểm thi" ||
          text.includes("điểm bằng số")
        ) {
          foundScore = colNumberToLetter(colNumber);
        }
      });
      if (foundSbd || foundScore) {
        detectedHeaderRow = r;
        if (foundSbd && !preferredSbdCol) detectedSbdCol = foundSbd;
        if (foundScore && !preferredScoreCol) detectedScoreCol = foundScore;
        break;
      }
    }
  }

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
  const maxCol = Math.max(worksheet.columnCount, 10);
  const headers: ExcelHeaderCol[] = [];

  for (let c = 1; c <= maxCol; c++) {
    const colLetter = colNumberToLetter(c);
    const rawLabel = String(extractCellValue(headerRowObj.getCell(c).value) ?? "")
      .replace(/[\r\n]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const label = rawLabel === "ST T" ? "STT" : rawLabel;
    const colWidth = worksheet.getColumn(c).width || 12;
    if (label || c <= 10) {
      headers.push({
        colLetter,
        colNumber: c,
        label: label || `Cột ${colLetter}`,
        width: Math.round(colWidth * 10) / 10,
      });
    }
  }

  const sbdColNum = colLetterToNumber(detectedSbdCol);
  const scoreColNum = colLetterToNumber(detectedScoreCol);
  const studentRows: ExcelStudentRow[] = [];

  for (let r = detectedHeaderRow + 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const sbdRaw = extractCellValue(row.getCell(sbdColNum).value);
    const sbdStr = sbdRaw !== null ? String(sbdRaw).trim() : "";

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
      currentScore,
      cells,
      fontName,
      hasBorder,
    });
  }

  const mergesCount =
    (worksheet.model && Array.isArray(worksheet.model.merges)
      ? worksheet.model.merges.length
      : 0) || 0;

  return {
    fileName: fileName || "Bang_Diem_Thi_Van_6_THCS_My_Xa.xlsx",
    fileBase64,
    availableSheets,
    currentSheetIndex: sheetIdx,
    sheetName: worksheet.name,
    totalRows: worksheet.rowCount,
    mergedCellsCount: mergesCount,
    detectedHeaderRow,
    detectedSbdCol,
    detectedScoreCol,
    titleRows,
    headers,
    studentRows,
  };
}

/**
 * 100% Client-Side Non-Destructive Excel Updater (Preserves all formatting, borders, merges, fonts)
 */
export async function exportUpdatedExcelClientSide(params: {
  fileBase64: string;
  fileName: string;
  headerRow: number;
  sbdCol: string;
  scoreCol: string;
  sheetIndex?: number;
  ocrRecords: Array<{ sbd: string; score: number }>;
}) {
  const { fileBase64, fileName, headerRow, sbdCol, scoreCol, sheetIndex, ocrRecords } =
    params;

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

  const ocrMap: Record<string, number> = {};
  for (const rec of ocrRecords || []) {
    const norm = normalizeSbdClient(rec.sbd);
    const canon = canonicalSbdKey(rec.sbd);
    const val = parseVietnameseScoreClient(rec.score);
    if (val !== null) {
      if (norm) ocrMap[norm] = val;
      if (canon) ocrMap[canon] = val;
    }
  }

  const sbdColNum = colLetterToNumber(sbdCol || "G");
  const scoreColNum = colLetterToNumber(scoreCol || "I");
  const startRow = (headerRow || 5) + 1;

  let matchedCount = 0;
  let untouchedCount = 0;
  const matchedDetails: Array<{ rowNumber: number; sbd: string; score: number }> = [];

  for (let r = startRow; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const rawSbd = extractCellValue(row.getCell(sbdColNum).value);
    const normSbd = normalizeSbdClient(rawSbd);
    const canonSbd = canonicalSbdKey(rawSbd);
    if (!normSbd && !canonSbd) continue;

    const matchedVal =
      (normSbd && Object.prototype.hasOwnProperty.call(ocrMap, normSbd)
        ? ocrMap[normSbd]
        : undefined) ??
      (canonSbd && Object.prototype.hasOwnProperty.call(ocrMap, canonSbd)
        ? ocrMap[canonSbd]
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
    `[Kiểm định Cấu trúc Bảng tính Sau Khi So Khớp SBD]`,
    ` - Trang tính đã xử lý: ${worksheet.name}`,
    ` - Số vùng gộp ô (Merged Cells) được bảo toàn: ${mergesCount}`,
    ` - Số ô cột ${scoreCol} (Điểm) đã ghi kiểu Float/Number: ${matchedCount}`,
    ...matchedDetails
      .slice(0, 10)
      .map((d) => `   + Hàng ${d.rowNumber} | SBD ${d.sbd} -> Ô ${scoreCol}${d.rowNumber} = ${d.score}`),
  ];

  const entries = Object.entries(ocrMap)
    .map(([k, v]) => `    "${k}": ${v},`)
    .join("\n");

  const openpyxlCode = `import re
import openpyxl

def normalize_sbd(value) -> str:
    """Chuẩn hóa mã Số Báo Danh (loại bỏ khoảng trắng thừa, viết hoa)."""
    if value is None:
        return ""
    return re.sub(r"\\s+", "", str(value).strip()).replace(",", ".").upper()

def update_exam_scores(input_excel_path: str, output_excel_path: str):
    # 1. Dữ liệu (SBD -> Điểm số Float) đã trích xuất từ Hình 1 (PHIẾU CHẤM BÀI THI)
    ocr_scores: dict[str, float] = {
${entries || "    # Chưa có dữ liệu điểm"}
    }

    # 2. Mở file Excel nguồn (Hình 2) — giữ nguyên hoàn toàn định dạng, merge cells, borders, font
    wb = openpyxl.load_workbook(input_excel_path, data_only=False)
    sheet_names = wb.sheetnames
    ws = wb[sheet_names[${targetIdx}]] if len(sheet_names) > ${targetIdx} else wb.active

    HEADER_ROW = ${headerRow}
    START_ROW = ${startRow}
    SBD_COL = "${sbdCol}"
    SCORE_COL = "${scoreCol}"

    matched_count = 0

    # 3. Quét danh sách SBD ở cột ${sbdCol} từ hàng ${startRow} trở đi và khớp điểm vào cột ${scoreCol} (Hình 3)
    for row_idx in range(START_ROW, ws.max_row + 1):
        raw_sbd = ws[f"{SBD_COL}{row_idx}"].value
        norm_sbd = normalize_sbd(raw_sbd)

        if norm_sbd and norm_sbd in ocr_scores:
            # Chỉ ghi đè giá trị kiểu Float/Number vào đúng ô Điểm (Cột ${scoreCol})
            ws[f"{SCORE_COL}{row_idx}"].value = float(ocr_scores[norm_sbd])
            matched_count += 1

    wb.save(output_excel_path)
    print(f"Hoàn tất! Đã khớp SBD và điền điểm {matched_count} học sinh vào cột {SCORE_COL}.")

if __name__ == "__main__":
    update_exam_scores("${fileName}", "Da_Nhap_Diem_${fileName}")
`;

  return {
    updatedExcelBase64,
    outputFileName: `Da_Nhap_Diem_${fileName || "Bang_Diem_Van_6_THCS_My_Xa.xlsx"}`,
    matchedCount,
    untouchedCount,
    matchedDetails,
    pythonAuditLog: auditLines.join("\n"),
    openpyxlCode,
  };
}

/**
 * Generates the exact Excel workbook from Hình 2:
 * SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH NINH BÌNH - TRƯỜNG THCS MỸ XÁ - BẢNG ĐIỂM THI MÔN: VĂN 6
 */
export async function generateSampleExcelClientSide(): Promise<ExcelSheetData> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Trường THCS Mỹ Xá";
  wb.created = new Date();

  const ws = wb.addWorksheet("Bang_Diem_Van_6", {
    views: [{ showGridLines: true }],
  });

  ws.columns = [
    { key: "stt", width: 6 },        // A: STT
    { key: "hoten", width: 24 },     // B: Họ và tên
    { key: "ngaysinh", width: 16 },  // C: Ngày sinh
    { key: "gioi", width: 9 },       // D: Giới
    { key: "phongthi", width: 10 },  // E: Phòng thi
    { key: "diadiem", width: 14 },   // F: Địa điểm thi
    { key: "sbd", width: 16 },       // G: Số BD
    { key: "made", width: 10 },      // H: Mã đề
    { key: "diem", width: 14 },      // I: Điểm
    { key: "ghichu", width: 14 },    // J: Ghi chú
  ];

  // Row 1: SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH NINH BÌNH
  ws.mergeCells("A1:D1");
  ws.getCell("A1").value = "SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH NINH BÌNH";
  ws.getCell("A1").font = { name: "Times New Roman", size: 12, bold: false };
  ws.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };

  // Row 2: TRƯỜNG THCS MỸ XÁ
  ws.mergeCells("A2:D2");
  ws.getCell("A2").value = "TRƯỜNG THCS MỸ XÁ";
  ws.getCell("A2").font = { name: "Times New Roman", size: 12, bold: true };
  ws.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  // Row 3: BẢNG ĐIỂM THI MÔN: VĂN 6
  ws.mergeCells("A3:J3");
  ws.getCell("A3").value = "BẢNG ĐIỂM THI MÔN: VĂN 6";
  ws.getCell("A3").font = { name: "Times New Roman", size: 13, bold: true };
  ws.getCell("A3").alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(3).height = 24;

  // Row 5: Table Header matching Hình 2
  const headerTitles = [
    "STT",
    "Họ và tên",
    "Ngày sinh",
    "Giới",
    "Phòng thi",
    "Địa điểm thi",
    "Số BD",
    "Mã đề",
    "Điểm",
    "Ghi chú",
  ];
  const headerRow = ws.getRow(5);
  headerRow.height = 30;
  headerTitles.forEach((title, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = title;
    cell.font = { name: "Times New Roman", size: 11, bold: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE5E7EB" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
  });

  // Exact students from Hình 2 (plus K6.006, K6.007 from Hình 1)
  const sampleStudents = [
    { stt: 1, hoten: "Đan Trường An", ngaysinh: "20/08/2015", gioi: "Nam", phong: 1, diadiem: "", sbd: "K6.001", made: "", ghichu: "6A4" },
    { stt: 2, hoten: "Đỗ Huy Khiết An", ngaysinh: "21/05/2015", gioi: "Nam", phong: 1, diadiem: "", sbd: "K6.002", made: "", ghichu: "6A3" },
    { stt: 3, hoten: "Đỗ Phúc Lâm An", ngaysinh: "14/07/2015", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K6.003", made: "", ghichu: "6A5" },
    { stt: 4, hoten: "Hoàng An An", ngaysinh: "16/12/2015", gioi: "Nam", phong: 1, diadiem: "", sbd: "K6.004", made: "", ghichu: "6A5" },
    { stt: 5, hoten: "Hoàng Khánh An", ngaysinh: "17/12/2015", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K6.005", made: "", ghichu: "6A2" },
    { stt: 6, hoten: "Lê Minh Anh", ngaysinh: "05/03/2015", gioi: "Nữ", phong: 1, diadiem: "", sbd: "K6.006", made: "", ghichu: "6A1" },
    { stt: 7, hoten: "Nguyễn Tuấn Anh", ngaysinh: "19/09/2015", gioi: "Nam", phong: 1, diadiem: "", sbd: "K6.007", made: "", ghichu: "6A2" },
  ];

  sampleStudents.forEach((student, index) => {
    const rowNum = 6 + index;
    const row = ws.getRow(rowNum);
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
      null, // Cột I: Điểm (Để trống như Hình 2, chờ điền tự động thành Hình 3)
      student.ghichu,
    ];
    values.forEach((val, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = val;
      cell.font = {
        name: "Times New Roman",
        size: 11,
        bold: false,
      };
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

  const buffer = await wb.xlsx.writeBuffer();
  const base64 = arrayBufferToBase64(buffer);
  return parseExcelClientSide({
    fileBase64: base64,
    fileName: "Bang_Diem_Thi_Van_6_THCS_My_Xa.xlsx",
    preferredHeaderRow: 5,
    preferredSbdCol: "G",
    preferredScoreCol: "I",
  });
}
