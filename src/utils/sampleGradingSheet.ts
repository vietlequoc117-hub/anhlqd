export interface OcrRecord {
  id: string;
  sbd: string;
  rawSbd: string;
  score: number;
  rawScoreText: string;
  scoreInWords: string;
  confidence: "CAO" | "TRUNG_BINH" | "CAN_KIEM_TRA";
  notes: string;
  sourceImage: string;
}

export interface ExcelHeaderCol {
  colLetter: string;
  colNumber: number;
  label: string;
  width: number;
}

export interface ExcelStudentRow {
  rowNumber: number;
  sbd: string;
  normalizedSbd: string;
  currentScore: number | string | null;
  cells: Record<string, string | number | null>;
  fontName?: string;
  hasBorder: boolean;
}

export interface ExcelSheetData {
  fileName: string;
  fileBase64: string;
  sheetName: string;
  currentSheetIndex?: number;
  availableSheets?: Array<{ index: number; name: string; rowCount: number }>;
  totalRows: number;
  mergedCellsCount: number;
  detectedHeaderRow: number;
  detectedSbdCol: string;
  detectedScoreCol: string;
  titleRows: Array<{ rowNumber: number; texts: string[] }>;
  headers: ExcelHeaderCol[];
  studentRows: ExcelStudentRow[];
}

export function normalizeSbdClient(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  return String(raw)
    .trim()
    .replace(/\s+/g, "")
    .replace(/,/g, ".")
    .toUpperCase();
}

export function parseVietnameseScoreClient(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number" && !Number.isNaN(raw)) {
    return Math.round(raw * 100) / 100;
  }
  const str = String(raw)
    .trim()
    .replace(/,/g, ".")
    .replace(/[^0-9.]/g, "");
  if (!str) return null;
  const val = parseFloat(str);
  if (Number.isNaN(val) || val < 0 || val > 10) return null;
  return Math.round(val * 100) / 100;
}

/**
 * Generates a realistic high-resolution image of a Vietnamese "PHIẾU CHẤM BÀI THI"
 * with teacher handwriting (including crossed 7s, open 4s, comma decimals) for live OCR testing.
 */
export function generateSampleGradingSheetImage(): {
  name: string;
  mimeType: string;
  dataUrl: string;
} {
  const canvas = document.createElement("canvas");
  canvas.width = 1000;
  canvas.height = 1220;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return { name: "Phieu_Cham_Bai_Thi_P01.png", mimeType: "image/png", dataUrl: "" };
  }

  // Warm paper background
  ctx.fillStyle = "#FCFBF7";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Subtle paper grain border
  ctx.strokeStyle = "#CBD5E1";
  ctx.lineWidth = 2;
  ctx.strokeRect(24, 24, canvas.width - 48, canvas.height - 48);

  // Header text (printed style)
  ctx.fillStyle = "#0F172A";
  ctx.font = "600 17px 'Plus Jakarta Sans', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("PHÒNG GD&ĐT QUẬN CẦU GIẤY", 250, 72);
  ctx.font = "700 18px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("TRƯỜNG THCS NGUYỄN DU", 250, 98);

  ctx.font = "700 17px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 720, 72);
  ctx.font = "500 16px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("Độc lập - Tự do - Hạnh phúc", 720, 98);

  // Main title
  ctx.font = "700 28px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("PHIẾU CHẤM BÀI THI", 500, 160);
  ctx.font = "500 17px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("Môn chấm: Toán 6   —   Túi bài số: 01   —   Phòng thi: P.01", 500, 192);

  // Table coordinates
  const tableLeft = 60;
  const tableTop = 230;
  const colWidths = [80, 230, 210, 240, 120]; // STT, SBD, ĐIỂM BẰNG SỐ, ĐIỂM BẰNG CHỮ, GHI CHÚ
  const rowHeight = 68;
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);

  // Table header background
  ctx.fillStyle = "#F1F5F9";
  ctx.fillRect(tableLeft, tableTop, totalWidth, rowHeight);

  // Draw table headers
  const headers = ["TT", "SỐ BÁO DANH (SBD)", "ĐIỂM BẰNG SỐ", "ĐIỂM BẰNG CHỮ", "KÝ TÊN"];
  ctx.fillStyle = "#0F172A";
  ctx.font = "700 16px 'Plus Jakarta Sans', sans-serif";

  let curX = tableLeft;
  headers.forEach((h, idx) => {
    ctx.textAlign = "center";
    ctx.fillText(h, curX + colWidths[idx] / 2, tableTop + 40);
    curX += colWidths[idx];
  });

  const rowsData = [
    { tt: "1", sbd: "K6.001", rawScore: "8,5", words: "Tám phẩy năm" },
    { tt: "2", sbd: "K6.002", rawScore: "6,7", words: "Sáu phẩy bảy" },
    { tt: "3", sbd: "K6.003", rawScore: "7,4", words: "Bảy phẩy bốn" },
    { tt: "4", sbd: "K6.004", rawScore: "9,0", words: "Chín điểm" },
    { tt: "5", sbd: "K6.005", rawScore: "7,5", words: "Bảy phẩy năm" },
    { tt: "6", sbd: "K6.006", rawScore: "10", words: "Mười điểm" },
    { tt: "7", sbd: "K6.007", rawScore: "5,6", words: "Năm phẩy sáu" },
    { tt: "8", sbd: "K6.008", rawScore: "8,0", words: "Tám điểm" },
    { tt: "9", sbd: "K6.009", rawScore: "6,0", words: "Sáu điểm" },
    { tt: "10", sbd: "K6.010", rawScore: "7,8", words: "Bảy phẩy tám" },
  ];

  // Draw grid lines
  ctx.strokeStyle = "#334155";
  ctx.lineWidth = 1.5;
  for (let r = 0; r <= rowsData.length + 1; r++) {
    const y = tableTop + r * rowHeight;
    ctx.beginPath();
    ctx.moveTo(tableLeft, y);
    ctx.lineTo(tableLeft + totalWidth, y);
    ctx.stroke();
  }

  let xLine = tableLeft;
  for (let c = 0; c <= colWidths.length; c++) {
    ctx.beginPath();
    ctx.moveTo(xLine, tableTop);
    ctx.lineTo(xLine, tableTop + (rowsData.length + 1) * rowHeight);
    ctx.stroke();
    if (c < colWidths.length) xLine += colWidths[c];
  }

  // Fill rows with blue ballpoint ink appearance for teacher handwriting
  rowsData.forEach((row, rIdx) => {
    const rowY = tableTop + (rIdx + 1) * rowHeight;
    const baselineY = rowY + 43;

    // Printed TT
    ctx.fillStyle = "#334155";
    ctx.font = "500 17px 'IBM Plex Mono', monospace";
    ctx.textAlign = "center";
    ctx.fillText(row.tt, tableLeft + colWidths[0] / 2, baselineY);

    // Handwritten SBD (blue ballpoint pen)
    ctx.fillStyle = "#1E3A8A";
    ctx.font = "600 22px 'IBM Plex Mono', monospace";
    ctx.fillText(row.sbd, tableLeft + colWidths[0] + colWidths[1] / 2, baselineY);

    // Handwritten Score in numbers (with comma decimal)
    ctx.fillStyle = "#1D4ED8";
    ctx.font = "700 26px 'IBM Plex Mono', monospace";
    const scoreX = tableLeft + colWidths[0] + colWidths[1] + colWidths[2] / 2;
    ctx.fillText(row.rawScore, scoreX, baselineY);

    // Handwritten Score in words
    ctx.fillStyle = "#1E3A8A";
    ctx.font = "italic 600 20px 'Plus Jakarta Sans', sans-serif";
    const wordsX = tableLeft + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] / 2;
    ctx.fillText(row.words, wordsX, baselineY);

    // Signature initials
    ctx.fillStyle = "#475569";
    ctx.font = "italic 16px 'Plus Jakarta Sans', sans-serif";
    const signX = tableLeft + colWidths[0] + colWidths[1] + colWidths[2] + colWidths[3] + colWidths[4] / 2;
    ctx.fillText("GV chấm", signX, baselineY);
  });

  // Footer signature section
  const footerY = tableTop + (rowsData.length + 1) * rowHeight + 60;
  ctx.fillStyle = "#0F172A";
  ctx.font = "600 16px 'Plus Jakarta Sans', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("GIÁM KHẢO 1", 260, footerY);
  ctx.fillText("GIÁM KHẢO 2", 740, footerY);

  ctx.font = "italic 15px 'Plus Jakarta Sans', sans-serif";
  ctx.fillStyle = "#475569";
  ctx.fillText("(Ký và ghi rõ họ tên)", 260, footerY + 24);
  ctx.fillText("(Ký và ghi rõ họ tên)", 740, footerY + 24);

  return {
    name: "Phieu_Cham_Bai_Thi_Toan6_Tui01.png",
    mimeType: "image/png",
    dataUrl: canvas.toDataURL("image/png"),
  };
}
