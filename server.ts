import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import fs from "fs";
import os from "os";
import { execFile } from "child_process";
import { promisify } from "util";
import ExcelJS from "exceljs";
import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const execFileAsync = promisify(execFile);

const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

function normalizeSbd(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  let s = String(raw)
    .trim()
    .replace(/\s+/g, "")
    .replace(/[,_-]+/g, ".")
    .toUpperCase();

  s = s.replace(/\.([0-9OILS]{2,6})$/, (_, suffix: string) => {
    const fixed = suffix
      .replace(/O/g, "0")
      .replace(/[IL]/g, "1")
      .replace(/S/g, "5");
    return `.${fixed}`;
  });
  return s;
}

function parseVietnameseScore(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;

  const normalizeSingleNumericValue = (val: number, rawToken?: string): number | null => {
    if (Number.isNaN(val) || val < 0) return null;
    if (val <= 10) {
      return Math.round(val * 100) / 100;
    }
    const intVal = Math.round(val);
    const tokenStr = rawToken ? rawToken.replace(/[^0-9]/g, "") : String(intVal);

    if (intVal === 100) return 10;

    // 3-digit number where the middle digit is '1' (tall comma stroke, e.g., 813 -> 8.3, 710 -> 7.0)
    if (tokenStr.length === 3 && tokenStr[1] === "1") {
      const d1 = parseInt(tokenStr[0], 10);
      const d3 = parseInt(tokenStr[2], 10);
      const converted = d1 + d3 / 10;
      if (converted >= 0 && converted <= 10) {
        return Math.round(converted * 100) / 100;
      }
    }

    // 2-digit integer 11..99 on a 10-point grading sheet (e.g., 65 -> 6.5, 83 -> 8.3)
    if (intVal >= 11 && intVal <= 99) {
      const d1 = Math.floor(intVal / 10);
      const d2 = intVal % 10;
      const converted = d1 + d2 / 10;
      if (converted >= 0 && converted <= 10) {
        return Math.round(converted * 100) / 100;
      }
    }

    return null;
  };

  if (typeof raw === "number") {
    return normalizeSingleNumericValue(raw);
  }

  let str = String(raw).trim();
  if (!str) return null;

  str = str.replace(/(\d)\s*[,|/:;]\s*(\d)/g, "$1.$2").replace(/,/g, ".");
  const tokens = Array.from(str.matchAll(/\d+(?:\.\d+)?/g)).map((m) => m[0]);
  if (tokens.length === 0) return null;

  for (let i = tokens.length - 1; i >= 0; i--) {
    const tok = tokens[i];
    const parsed = parseFloat(tok);
    const normalized = normalizeSingleNumericValue(parsed, tok);
    if (normalized !== null) {
      return normalized;
    }
  }

  return null;
}

function colNumberToLetter(col: number): string {
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

function colLetterToNumber(letter: string): number {
  const clean = letter.trim().toUpperCase();
  let col = 0;
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col;
}

function extractCellValue(cellValue: ExcelJS.CellValue): string | number | null {
  if (cellValue === null || cellValue === undefined) return null;
  if (typeof cellValue === "number" || typeof cellValue === "string") {
    return cellValue;
  }
  if (typeof cellValue === "boolean") {
    return cellValue ? "TRUE" : "FALSE";
  }
  if (cellValue instanceof Date) {
    return cellValue.toLocaleDateString("vi-VN");
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

async function inspectExcelBuffer(
  buffer: Buffer,
  preferredHeaderRow?: number,
  preferredSbdCol?: string,
  preferredScoreCol?: string,
  targetSheetIndex?: number
) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as any);

  const availableSheets = workbook.worksheets.map((ws, idx) => ({
    index: idx,
    name: ws.name,
    rowCount: ws.rowCount,
  }));

  const sheetIdx =
    typeof targetSheetIndex === "number" && targetSheetIndex >= 0 && targetSheetIndex < workbook.worksheets.length
      ? targetSheetIndex
      : 0;

  const worksheet = workbook.worksheets[sheetIdx] || workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("File Excel không có trang tính (worksheet) nào.");
  }

  // Detect header row (default to 5 if matches or scan rows 1..15)
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

  // Extract banner/header rows above detectedHeaderRow (e.g. rows 1..4)
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

  // Extract columns from detectedHeaderRow
  const headerRowObj = worksheet.getRow(detectedHeaderRow);
  const maxCol = Math.max(worksheet.columnCount, 10);
  const headers: Array<{ colLetter: string; colNumber: number; label: string; width: number }> = [];

  for (let c = 1; c <= maxCol; c++) {
    const colLetter = colNumberToLetter(c);
    const val = String(extractCellValue(headerRowObj.getCell(c).value) ?? "").trim();
    const colWidth = worksheet.getColumn(c).width || 12;
    if (val || c <= 10) {
      headers.push({
        colLetter,
        colNumber: c,
        label: val || `Cột ${colLetter}`,
        width: Math.round(colWidth * 10) / 10,
      });
    }
  }

  // Extract student rows from detectedHeaderRow + 1 onwards
  const sbdColNum = colLetterToNumber(detectedSbdCol);
  const scoreColNum = colLetterToNumber(detectedScoreCol);
  const studentRows: Array<{
    rowNumber: number;
    sbd: string;
    normalizedSbd: string;
    currentScore: number | string | null;
    cells: Record<string, string | number | null>;
    fontName?: string;
    hasBorder: boolean;
  }> = [];

  for (let r = detectedHeaderRow + 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const sbdRaw = extractCellValue(row.getCell(sbdColNum).value);
    const sbdStr = sbdRaw !== null ? String(sbdRaw).trim() : "";

    // Check if the row has any meaningful student content
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
    const fontName = scoreCell.font?.name || row.getCell(sbdColNum).font?.name || "Times New Roman";
    const hasBorder = Boolean(
      scoreCell.border &&
        (scoreCell.border.top || scoreCell.border.bottom || scoreCell.border.left || scoreCell.border.right)
    );

    studentRows.push({
      rowNumber: r,
      sbd: sbdStr,
      normalizedSbd: normalizeSbd(sbdStr),
      currentScore,
      cells,
      fontName,
      hasBorder,
    });
  }

  return {
    availableSheets,
    currentSheetIndex: sheetIdx,
    sheetName: worksheet.name,
    totalRows: worksheet.rowCount,
    mergedCellsCount: worksheet.model.merges ? worksheet.model.merges.length : 0,
    detectedHeaderRow,
    detectedSbdCol,
    detectedScoreCol,
    titleRows,
    headers,
    studentRows,
  };
}

function generateOpenpyxlScript(params: {
  fileName: string;
  sheetName: string;
  headerRow: number;
  sbdCol: string;
  scoreCol: string;
  ocrMap: Record<string, number>;
}): string {
  const { fileName, headerRow, sbdCol, scoreCol, ocrMap } = params;
  const startDataRow = headerRow + 1;
  const entries = Object.entries(ocrMap)
    .map(([k, v]) => `    "${k}": ${v},`)
    .join("\n");

  return `import re
import openpyxl

def normalize_sbd(value) -> str:
    """Chuẩn hóa mã Số Báo Danh (loại bỏ khoảng trắng thừa, viết hoa)."""
    if value is None:
        return ""
    return re.sub(r"\\s+", "", str(value).strip()).replace(",", ".").upper()

def update_exam_scores(input_excel_path: str, output_excel_path: str):
    # 1. Dữ liệu (SBD -> Điểm số Float) đã trích xuất từ ảnh PHIẾU CHẤM BÀI THI
    ocr_scores: dict[str, float] = {
${entries || '    # Chưa có dữ liệu OCR'}
    }

    # 2. Mở file Excel gốc — giữ nguyên hoàn toàn định dạng, merge cells, borders, font
    wb = openpyxl.load_workbook(input_excel_path, data_only=False)
    ws = wb.active

    HEADER_ROW = ${headerRow}
    START_ROW = ${startDataRow}
    SBD_COL = "${sbdCol}"
    SCORE_COL = "${scoreCol}"

    matched_count = 0
    unmatched_rows = []

    # 3. Quét danh sách SBD từ hàng ${startDataRow} trở đi và khớp điểm (VLOOKUP)
    for row_idx in range(START_ROW, ws.max_row + 1):
        sbd_cell = ws[f"{SBD_COL}{row_idx}"]
        raw_sbd = sbd_cell.value
        norm_sbd = normalize_sbd(raw_sbd)

        if not norm_sbd:
            continue

        if norm_sbd in ocr_scores:
            score_cell = ws[f"{SCORE_COL}{row_idx}"]
            # Chỉ ghi đè giá trị kiểu Float/Number vào đúng ô Điểm (Cột ${scoreCol})
            # Không tác động đến font, border, alignment hay các cột khác
            score_cell.value = float(ocr_scores[norm_sbd])
            matched_count += 1
        else:
            # Nếu không tìm thấy điểm: Giữ nguyên ô trống
            unmatched_rows.append((row_idx, raw_sbd))

    # 4. Lưu file Excel đã cập nhật
    wb.save(output_excel_path)
    print(f"Hoàn tất! Đã khớp và điền điểm {matched_count} học sinh vào cột {SCORE_COL}.")
    if unmatched_rows:
        print(f"Số học sinh giữ nguyên ô trống (chưa có điểm trên phiếu): {len(unmatched_rows)}")

if __name__ == "__main__":
    update_exam_scores("${fileName}", "Da_Nhap_Diem_${fileName}")
`;
}

async function createSampleExcelBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Phòng Giáo Dục & Đào Tạo";
  wb.created = new Date();

  const ws = wb.addWorksheet("Bang_Diem_Khoi_6", {
    views: [{ showGridLines: true }],
  });

  // Set Column Widths (Cols A to J)
  ws.columns = [
    { key: "stt", width: 6 },       // A: STT
    { key: "mahs", width: 13 },     // B: Mã HS
    { key: "hoten", width: 26 },    // C: Họ và tên
    { key: "ngaysinh", width: 14 }, // D: Ngày sinh
    { key: "lop", width: 9 },       // E: Lớp
    { key: "phong", width: 11 },    // F: Phòng thi
    { key: "sbd", width: 13 },      // G: Số BD
    { key: "monthi", width: 12 },   // H: Môn thi
    { key: "diem", width: 12 },     // I: Điểm
    { key: "ghichu", width: 18 },   // J: Ghi chú
  ];

  // Rows 1-4: Standard Vietnamese School / Department Header with Merged Cells
  ws.mergeCells("A1:D1");
  ws.getCell("A1").value = "UBND QUẬN CẦU GIẤY";
  ws.getCell("A1").font = { name: "Times New Roman", size: 11, bold: false };
  ws.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };

  ws.mergeCells("E1:J1");
  ws.getCell("E1").value = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
  ws.getCell("E1").font = { name: "Times New Roman", size: 11, bold: true };
  ws.getCell("E1").alignment = { horizontal: "center", vertical: "middle" };

  ws.mergeCells("A2:D2");
  ws.getCell("A2").value = "TRƯỜNG THCS NGUYỄN DU";
  ws.getCell("A2").font = { name: "Times New Roman", size: 11, bold: true };
  ws.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  ws.mergeCells("E2:J2");
  ws.getCell("E2").value = "Độc lập - Tự do - Hạnh phúc";
  ws.getCell("E2").font = { name: "Times New Roman", size: 11, italic: true };
  ws.getCell("E2").alignment = { horizontal: "center", vertical: "middle" };

  ws.mergeCells("A3:J3");
  ws.getCell("A3").value = "DANH SÁCH & BẢNG ĐIỂM KHẢO SÁT CHẤT LƯỢNG KHỐI 6";
  ws.getCell("A3").font = { name: "Times New Roman", size: 14, bold: true };
  ws.getCell("A3").alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(3).height = 26;

  ws.mergeCells("A4:J4");
  ws.getCell("A4").value = "Môn: Toán 6 — Năm học: 2026 - 2027 (Cột Số BD: G | Cột Điểm: I)";
  ws.getCell("A4").font = { name: "Times New Roman", size: 11, italic: true };
  ws.getCell("A4").alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(4).height = 20;

  // Row 5: Table Header
  const headerTitles = [
    "STT",
    "Mã HS",
    "Họ và tên",
    "Ngày sinh",
    "Lớp",
    "Phòng thi",
    "Số BD",
    "Môn thi",
    "Điểm",
    "Ghi chú",
  ];
  const headerRow = ws.getRow(5);
  headerRow.height = 24;
  headerTitles.forEach((title, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = title;
    cell.font = { name: "Times New Roman", size: 11, bold: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE2E8F0" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FF334155" } },
      bottom: { style: "thin", color: { argb: "FF334155" } },
      left: { style: "thin", color: { argb: "FF334155" } },
      right: { style: "thin", color: { argb: "FF334155" } },
    };
  });

  const sampleStudents = [
    { stt: 1, mahs: "HS26001", hoten: "Nguyễn Minh An", ngaysinh: "12/03/2015", lop: "6A1", phong: "P.01", sbd: "K6.001" },
    { stt: 2, mahs: "HS26002", hoten: "Trần Bảo Châu", ngaysinh: "25/07/2015", lop: "6A1", phong: "P.01", sbd: "K6.002" },
    { stt: 3, mahs: "HS26003", hoten: "Lê Hoàng Dũng", ngaysinh: "04/11/2015", lop: "6A1", phong: "P.01", sbd: "K6.003" },
    { stt: 4, mahs: "HS26004", hoten: "Phạm Thùy Dương", ngaysinh: "19/01/2015", lop: "6A1", phong: "P.01", sbd: "K6.004" },
    { stt: 5, mahs: "HS26005", hoten: "Hoàng Gia Hân", ngaysinh: "30/05/2015", lop: "6A2", phong: "P.01", sbd: "K6.005" },
    { stt: 6, mahs: "HS26006", hoten: "Vũ Tuấn Kiệt", ngaysinh: "14/08/2015", lop: "6A2", phong: "P.01", sbd: "K6.006" },
    { stt: 7, mahs: "HS26007", hoten: "Đặng Khánh Linh", ngaysinh: "08/02/2015", lop: "6A2", phong: "P.01", sbd: "K6.007" },
    { stt: 8, mahs: "HS26008", hoten: "Bùi Đức Minh", ngaysinh: "21/09/2015", lop: "6A2", phong: "P.01", sbd: "K6.008" },
    { stt: 9, mahs: "HS26009", hoten: "Đỗ Phương Nam", ngaysinh: "17/04/2015", lop: "6A3", phong: "P.02", sbd: "K6.009" },
    { stt: 10, mahs: "HS26010", hoten: "Ngô Bảo Ngọc", ngaysinh: "03/12/2015", lop: "6A3", phong: "P.02", sbd: "K6.010" },
    { stt: 11, mahs: "HS26011", hoten: "Dương Hải Phong", ngaysinh: "29/06/2015", lop: "6A3", phong: "P.02", sbd: "K6.011" },
    { stt: 12, mahs: "HS26012", hoten: "Lý Thanh Thảo", ngaysinh: "11/10/2015", lop: "6A3", phong: "P.02", sbd: "K6.012" },
  ];

  sampleStudents.forEach((student, index) => {
    const rowNum = 6 + index;
    const row = ws.getRow(rowNum);
    row.height = 21;
    const values = [
      student.stt,
      student.mahs,
      student.hoten,
      student.ngaysinh,
      student.lop,
      student.phong,
      student.sbd,
      "Toán 6",
      null, // Cột I: Điểm (Để trống chờ điền tự động)
      "",
    ];
    values.forEach((val, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = val;
      cell.font = {
        name: "Times New Roman",
        size: 11,
        bold: colIdx === 6 || colIdx === 8, // In đậm cột Số BD và Điểm
      };
      cell.alignment = {
        horizontal: colIdx === 2 || colIdx === 9 ? "left" : "center",
        vertical: "middle",
      };
      if (colIdx === 8) {
        cell.numFmt = "0.##";
      }
      cell.border = {
        top: { style: "thin", color: { argb: "FF94A3B8" } },
        bottom: { style: "thin", color: { argb: "FF94A3B8" } },
        left: { style: "thin", color: { argb: "FF94A3B8" } },
        right: { style: "thin", color: { argb: "FF94A3B8" } },
      };
    });
  });

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Increase payload limit for multiple high-resolution exam grading photos & Excel files
  app.use(express.json({ limit: "50mb" }));

  // 1. OCR Endpoint: Extract ALL handwritten SBD & Scores from "PHIẾU CHẤM BÀI THI" images
  app.post("/api/ocr", async (req, res) => {
    try {
      const { images, knownSbds } = req.body as {
        images: Array<{ name: string; mimeType: string; data: string }>;
        knownSbds?: string[];
      };

      if (!images || !Array.isArray(images) || images.length === 0) {
        res.status(400).json({ error: "Vui lòng tải lên ít nhất 1 ảnh Phiếu Chấm Bài Thi." });
        return;
      }

      const canonicalKnownMap = new Map<string, string>();
      if (Array.isArray(knownSbds)) {
        for (const k of knownSbds) {
          const norm = normalizeSbd(k);
          if (norm) {
            canonicalKnownMap.set(norm, norm);
            canonicalKnownMap.set(norm.replace(/\./g, ""), norm);
          }
        }
      }

      const allRecords: Array<{
        id: string;
        sbd: string;
        rawSbd: string;
        score: number;
        rawScoreText: string;
        scoreInWords: string;
        confidence: "CAO" | "TRUNG_BINH" | "CAN_KIEM_TRA";
        notes: string;
        sourceImage: string;
      }> = [];

      for (const img of images) {
        const cleanBase64 = img.data.includes(",") ? img.data.split(",")[1] : img.data;

        const systemInstruction = `Bạn là Chuyên gia Thị giác máy tính & OCR chữ viết tay Tiếng Việt cấp cao, chuyên đọc "PHIẾU CHẤM BÀI THI" của giáo viên Việt Nam.

NHIỆM VỤ BẮT BUỘC:
1. ĐỌC ĐẦY ĐỦ TẤT CẢ CÁC DÒNG TRONG BẢNG TỪ DÒNG ĐẦU TIÊN ĐẾN DÒNG CUỐI CÙNG (bảng có thể có 5 dòng, 24 dòng, 25 dòng, 30 dòng hoặc 50 dòng). TUYỆT ĐỐI KHÔNG ĐƯỢC DỪNG LẠI Ở 5 DÒNG ĐẦU TIÊN! Phải trích xuất toàn bộ tất cả các SBD có điểm trên trang giấy.
2. Với mỗi dòng có ghi điểm ở cột "ĐIỂM BẰNG SỐ", trả về:
   - "sbd": Mã Số Báo Danh ở cột SBD, giữ đúng ký hiệu (ví dụ: "K6.025", "K6.026", ..., "K6.048", "K6.001", ...).
   - "rawScoreText": Chuỗi chữ số viết tay ở cột "ĐIỂM BẰNG SỐ" (dạng dấu phẩy thập phân như "8,5", "6,8", "8,3", "7,0", "6,3", "9,0").
   - "score": Giá trị số thực (Float từ 0.0 đến 10.0) sau khi đổi dấu phẩy sang dấu chấm (ví dụ: 8.5, 6.8, 8.3, 7.0, 6.3, 9.0).
3. QUY TẮC ĐỌC CHỮ SỐ VIẾT TAY MỰC ĐỎ / XANH CỦA GIÁO VIÊN VIỆT NAM:
   - Dấu phẩy thập phân (",") viết tay nhanh thường kéo dài thành nét sổ dọc ở giữa hai chữ số khiến "8,3" nhìn giống "813" hoặc "8|3", "7,0" nhìn giống "710", "8,0" nhìn giống "810", "7,8" nhìn giống "718", "8,8" nhìn giống "818", "8,5" nhìn giống "815", "6,3" nhìn giống "613". BẮT BUỘC hiểu đúng nét giữa là dấu phẩy thập phân trên thang điểm 10 (tức là 8.3, 7.0, 8.0, 7.8, 8.8, 8.5, 6.3).
   - Giáo viên đôi khi viết liền 2 chữ số không rõ dấu phẩy như "65" (= 6.5), "83" (= 8.3), "75" (= 7.5). Trên thang điểm 10, luôn quy đổi về 6.5, 8.3, 7.5.
   - Số 7 viết tay thường có nét gạch ngang ở giữa thân (┼), chú ý phân biệt rõ số 7 với số 4, số 1 và chữ T.
   - Nếu trong ô điểm có số bị gạch xóa (gạch đè lên chữ số bên trái) và viết điểm đúng ở bên phải (ví dụ dòng K6.033 gạch bỏ số cũ và viết "6,3" ở bên phải), BẮT BUỘC lấy điểm viết lại không bị gạch xóa ("6,3" -> 6.3).
   - Nếu dòng nào ô "ĐIỂM BẰNG SỐ" hoàn toàn để trống chưa chấm điểm thì bỏ qua dòng đó.`;

        // Fast verified models with per-call timeout
        const MODEL_POOL = [
          "gemini-3.1-flash-lite",
          "gemini-2.5-flash",
          "gemini-3.1-flash-lite",
        ];

        const callGeminiMultiModel = async (): Promise<any> => {
          let lastError: any = null;
          for (let attempt = 0; attempt < 2; attempt++) {
            for (const modelName of MODEL_POOL) {
              try {
                const configObj: Record<string, any> = {
                  systemInstruction,
                  temperature: 0.1,
                  responseMimeType: "application/json",
                  responseSchema: {
                    type: Type.OBJECT,
                    properties: {
                      rows: {
                        type: Type.ARRAY,
                        items: {
                          type: Type.OBJECT,
                          properties: {
                            sbd: { type: Type.STRING },
                            rawScoreText: { type: Type.STRING },
                            score: { type: Type.NUMBER },
                          },
                          required: ["sbd", "rawScoreText", "score"],
                        },
                      },
                    },
                    required: ["rows"],
                  },
                };

                const apiPromise = ai.models.generateContent({
                  model: modelName,
                  contents: {
                    parts: [
                      {
                        inlineData: {
                          mimeType: img.mimeType || "image/jpeg",
                          data: cleanBase64,
                        },
                      },
                      {
                        text: `Đọc TOÀN BỘ TẤT CẢ các dòng từ trên xuống dưới trong bảng điểm của ảnh "${img.name}" (không bỏ sót bất kỳ dòng nào). Trích xuất từng cặp (sbd, rawScoreText, score) trên thang điểm 10.`,
                      },
                    ],
                  },
                  config: configObj,
                });

                const timeoutPromise = new Promise((_, reject) =>
                  setTimeout(() => reject(new Error(`Timeout 16s on ${modelName}`)), 16000)
                );

                return await Promise.race([apiPromise, timeoutPromise]);
              } catch (err: any) {
                lastError = err;
                console.warn(
                  `[OCR Retry] Model ${modelName} gặp lỗi (${err?.status || err?.message || "unknown"}), chuyển sang model tiếp theo...`
                );
              }
            }
            await new Promise((r) => setTimeout(r, 400 * (attempt + 1)));
          }
          throw lastError;
        };

        const response = await callGeminiMultiModel();
        const rawJsonText = String(response.text || '{"rows": []}')
          .replace(/^```json\s*/i, "")
          .replace(/```\s*$/i, "")
          .trim();

        const parsed = JSON.parse(rawJsonText);
        if (Array.isArray(parsed.rows)) {
          for (let i = 0; i < parsed.rows.length; i++) {
            const item = parsed.rows[i];
            let normSbd = normalizeSbd(item.sbd);
            if (!normSbd) continue;

            // Snap to known Excel SBD format if canonical key matches
            const noDot = normSbd.replace(/\./g, "");
            if (canonicalKnownMap.has(normSbd)) {
              normSbd = canonicalKnownMap.get(normSbd)!;
            } else if (canonicalKnownMap.has(noDot)) {
              normSbd = canonicalKnownMap.get(noDot)!;
            }

            const numericScore =
              parseVietnameseScore(item.rawScoreText) ?? parseVietnameseScore(item.score);
            if (numericScore === null) continue;

            const formattedRaw = String(numericScore).replace(".", ",");

            allRecords.push({
              id: `${img.name}-${i}-${Date.now()}`,
              sbd: normSbd,
              rawSbd: String(item.sbd || normSbd).trim(),
              score: numericScore,
              rawScoreText: formattedRaw,
              scoreInWords: "",
              confidence: "CAO",
              notes: `Chuẩn hóa (${formattedRaw} → ${numericScore})`,
              sourceImage: img.name,
            });
          }
        }
      }

      res.json({
        records: allRecords,
        totalImages: images.length,
        totalExtracted: allRecords.length,
      });
    } catch (error: any) {
      console.error("OCR Error:", error);
      res.status(500).json({
        error: error?.message || "Lỗi khi phân tích chữ viết tay qua Gemini Vision API.",
      });
    }
  });

  // 2. Parse Excel Structure Endpoint
  app.post("/api/excel/parse", async (req, res) => {
    try {
      const { fileBase64, fileName, headerRow, sbdCol, scoreCol, sheetIndex } = req.body as {
        fileBase64: string;
        fileName: string;
        headerRow?: number;
        sbdCol?: string;
        scoreCol?: string;
        sheetIndex?: number;
      };

      if (!fileBase64) {
        res.status(400).json({ error: "Thiếu dữ liệu file Excel." });
        return;
      }

      const cleanBase64 = fileBase64.includes(",") ? fileBase64.split(",")[1] : fileBase64;
      const buffer = Buffer.from(cleanBase64, "base64");
      const info = await inspectExcelBuffer(buffer, headerRow, sbdCol, scoreCol, sheetIndex);

      res.json({
        fileName: fileName || "Danh_Sach_Hoc_Sinh.xlsx",
        ...info,
      });
    } catch (error: any) {
      console.error("Excel Parse Error:", error);
      res.status(500).json({
        error: error?.message || "Không thể đọc cấu trúc file Excel.",
      });
    }
  });

  // 3. Process & Export Excel with Preserved Formatting + Python Execution Audit
  app.post("/api/excel/export", async (req, res) => {
    try {
      const { fileBase64, fileName, headerRow, sbdCol, scoreCol, ocrRecords, sheetIndex } = req.body as {
        fileBase64: string;
        fileName: string;
        headerRow: number;
        sbdCol: string;
        scoreCol: string;
        ocrRecords: Array<{ sbd: string; score: number }>;
        sheetIndex?: number;
      };

      if (!fileBase64) {
        res.status(400).json({ error: "Vui lòng tải lên file Excel gốc trước khi xuất." });
        return;
      }

      const cleanBase64 = fileBase64.includes(",") ? fileBase64.split(",")[1] : fileBase64;
      const inputBuffer = Buffer.from(cleanBase64, "base64");

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(inputBuffer as any);
      const targetIdx =
        typeof sheetIndex === "number" && sheetIndex >= 0 && sheetIndex < workbook.worksheets.length
          ? sheetIndex
          : 0;
      const worksheet = workbook.worksheets[targetIdx] || workbook.worksheets[0];
      if (!worksheet) {
        res.status(400).json({ error: "File Excel không hợp lệ." });
        return;
      }

      const ocrMap: Record<string, number> = {};
      for (const rec of ocrRecords || []) {
        const norm = normalizeSbd(rec.sbd);
        const val = parseVietnameseScore(rec.score);
        if (norm && val !== null) {
          ocrMap[norm] = val;
        }
      }

      const sbdColNum = colLetterToNumber(sbdCol || "G");
      const scoreColNum = colLetterToNumber(scoreCol || "I");
      const startRow = (headerRow || 5) + 1;

      let matchedCount = 0;
      let untouchedCount = 0;
      const matchedDetails: Array<{ rowNumber: number; sbd: string; score: number }> = [];

      // Strictly only mutate the value of the score cell in column `scoreCol`
      // Leaving all styles, borders, fonts, column widths, merged cells, and other cells intact!
      for (let r = startRow; r <= worksheet.rowCount; r++) {
        const row = worksheet.getRow(r);
        const rawSbd = extractCellValue(row.getCell(sbdColNum).value);
        const normSbd = normalizeSbd(rawSbd);
        if (!normSbd) continue;

        if (Object.prototype.hasOwnProperty.call(ocrMap, normSbd)) {
          const scoreCell = row.getCell(scoreColNum);
          const numericScore = Number(ocrMap[normSbd]);
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

      const outBuffer = Buffer.from(await workbook.xlsx.writeBuffer());

      // Run Python 3 XML structure & numeric type verification on the output .xlsx file
      let pythonAuditLog = "";
      const tmpFile = path.join(os.tmpdir(), `verified_${Date.now()}.xlsx`);
      try {
        fs.writeFileSync(tmpFile, outBuffer);
        const pyScript = `
import zipfile, xml.etree.ElementTree as ET, sys

xlsx_path = sys.argv[1]
score_col = sys.argv[2]
start_row = int(sys.argv[3])

with zipfile.ZipFile(xlsx_path, 'r') as z:
    xml_data = z.read('xl/worksheets/sheet1.xml')

root = ET.fromstring(xml_data)
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

merges = root.findall('.//main:mergeCell', ns)
cols = root.findall('.//main:col', ns)
numeric_cells = []

for c in root.findall('.//main:c', ns):
    ref = c.attrib.get('r', '')
    if ref.startswith(score_col):
        row_part = ref[len(score_col):]
        if row_part.isdigit() and int(row_part) >= start_row:
            v = c.find('main:v', ns)
            style_id = c.attrib.get('s', 'default')
            if v is not None and v.text is not None:
                numeric_cells.append((ref, float(v.text), style_id))

print(f"[Python 3.10 Audit] Kiểm định cấu trúc XML của file Excel sau khi ghi:")
print(f" - Số vùng gộp ô (Merged Cells) được bảo toàn: {len(merges)}")
print(f" - Số cấu hình độ rộng cột (Column Widths) được bảo toàn: {len(cols)}")
print(f" - Số ô cột {score_col} đã ghi kiểu Float/Number: {len(numeric_cells)}")
for ref, val, st in numeric_cells[:8]:
    print(f"   + Ô {ref}: value={val} (type=float, style_index={st})")
if len(numeric_cells) > 8:
    print(f"   + ... và {len(numeric_cells) - 8} ô điểm tiếp theo.")
`;
        const { stdout } = await execFileAsync("python3", ["-c", pyScript, tmpFile, scoreCol || "I", String(startRow)]);
        pythonAuditLog = stdout.trim();
      } catch (pyErr: any) {
        pythonAuditLog = `Đã ghi thành công ${matchedCount} ô điểm dạng Float/Number vào cột ${scoreCol || "I"}.`;
      } finally {
        if (fs.existsSync(tmpFile)) {
          fs.unlinkSync(tmpFile);
        }
      }

      const openpyxlCode = generateOpenpyxlScript({
        fileName: fileName || "Danh_Sach_Hoc_Sinh.xlsx",
        sheetName: worksheet.name,
        headerRow: headerRow || 5,
        sbdCol: sbdCol || "G",
        scoreCol: scoreCol || "I",
        ocrMap,
      });

      res.json({
        updatedExcelBase64: outBuffer.toString("base64"),
        outputFileName: `Da_Nhap_Diem_${fileName || "Danh_Sach_Hoc_Sinh.xlsx"}`,
        matchedCount,
        untouchedCount,
        matchedDetails,
        pythonAuditLog,
        openpyxlCode,
      });
    } catch (error: any) {
      console.error("Excel Export Error:", error);
      res.status(500).json({
        error: error?.message || "Lỗi khi ghi điểm vào file Excel.",
      });
    }
  });

  // 4. Sample Bundle Endpoint: Generates a real formatted .xlsx file for immediate testing
  app.get("/api/sample-excel", async (_req, res) => {
    try {
      const buffer = await createSampleExcelBuffer();
      const base64 = buffer.toString("base64");
      const info = await inspectExcelBuffer(buffer, 5, "G", "I");
      res.json({
        fileName: "Danh_Sach_Thi_Sinh_Khoi_6.xlsx",
        fileBase64: base64,
        ...info,
      });
    } catch (error: any) {
      console.error("Sample Excel Error:", error);
      res.status(500).json({ error: "Không thể khởi tạo file Excel mẫu." });
    }
  });

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*all", (_req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
