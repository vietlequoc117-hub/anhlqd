export interface OcrRecord {
  id: string;
  sbd: string;
  rawSbd: string;
  maPhach?: string;
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
  maPhach?: string;
  normalizedPhach?: string;
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
  detectedPhachCol?: string;
  detectedScoreCol: string;
  titleRows: Array<{ rowNumber: number; texts: string[] }>;
  headers: ExcelHeaderCol[];
  studentRows: ExcelStudentRow[];
  /** Cross-sheet SBD <-> Mã Phách pairs when loaded as File Excel 2 (Bảng Mã Phách) */
  allSheetsPhachPairs?: Array<{
    sheetName: string;
    rowNumber: number;
    sbd: string;
    normalizedSbd: string;
    maPhach: string;
    normalizedPhach: string;
  }>;
}

/**
 * Normalizes SBD or Mã Phách codes (e.g. "K9.001", "K6.025", "A40", "A 41", "A4O" -> "K9.001", "K6.025", "A40", "A41", "A40")
 */
export function normalizeSbdClient(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  let s = String(raw)
    .trim()
    .replace(/\s+/g, "")
    .replace(/[,_-]+/g, ".")
    .toUpperCase();

  // Fix common OCR confusions in the numeric suffix of SBD (e.g., K6.O25 -> K6.025, K9.0O1 -> K9.001)
  s = s.replace(/\.([0-9OILS]{2,6})$/, (_, suffix: string) => {
    const fixed = suffix
      .replace(/O/g, "0")
      .replace(/[IL]/g, "1")
      .replace(/S/g, "5");
    return `.${fixed}`;
  });

  // Also fix OCR confusions in compact Mã Phách codes like "A4O" -> "A40", "B0I" -> "B01"
  s = s.replace(/^([A-Z]{1,3})([0-9OILS]{1,5})$/, (_, prefix: string, digits: string) => {
    // Only convert if digits contains at least one digit or looks like a phách number
    if (/[0-9O]/.test(digits)) {
      const fixed = digits
        .replace(/O/g, "0")
        .replace(/[IL]/g, "1")
        .replace(/S/g, "5");
      return `${prefix}${fixed}`;
    }
    return `${prefix}${digits}`;
  });

  return s;
}

/**
 * Produces a canonical lookup key without dots/punctuation (e.g., "K9.001" -> "K9001", "A40" -> "A40")
 * so minor punctuation differences never prevent VLOOKUP matching.
 */
export function canonicalSbdKey(raw: unknown): string {
  return normalizeSbdClient(raw).replace(/[^A-Z0-9]/g, "");
}

/**
 * Parses a Vietnamese handwritten exam score on a 0.0 - 10.0 scale.
 * Handles:
 * - Standard comma/dot decimals: "8,5" -> 8.5, "8,50" -> 8.5, "7,0" -> 7.0, "8" -> 8
 * - Tall comma stroke read as '1' or '|': "813" -> 8.3, "710" -> 7.0, "810" -> 8.0, "718" -> 7.8, "613" -> 6.3, "8150" -> 8.5
 * - Fast 2-digit or 3-digit scores without visible comma: "65" -> 6.5, "83" -> 8.3, "850" -> 8.5, "725" -> 7.25
 * - Crossed-out corrections where the final score is written on the right: "6,8 6,3" -> 6.3
 */
export function parseVietnameseScoreClient(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;

  const normalizeSingleNumericValue = (val: number, rawToken?: string): number | null => {
    if (Number.isNaN(val) || val < 0) return null;
    if (val <= 10) {
      return Math.round(val * 100) / 100;
    }
    const intVal = Math.round(val);
    const tokenStr = rawToken ? rawToken.replace(/[^0-9]/g, "") : String(intVal);

    // 100 or 1000 -> 10.0
    if (intVal === 100 || intVal === 1000) return 10;

    // 4-digit number where the 2nd digit is '1' (tall comma stroke in 2-decimal score, e.g., 8150 -> 8.5, 7125 -> 7.25)
    if (tokenStr.length === 4 && tokenStr[1] === "1") {
      const d1 = parseInt(tokenStr[0], 10);
      const dec = parseInt(tokenStr.slice(2), 10);
      const converted = d1 + dec / 100;
      if (converted >= 0 && converted <= 10) {
        return Math.round(converted * 100) / 100;
      }
    }

    // 3-digit number where the middle digit is '1' (tall comma stroke, e.g., 813 -> 8.3, 710 -> 7.0)
    if (tokenStr.length === 3 && tokenStr[1] === "1") {
      const d1 = parseInt(tokenStr[0], 10);
      const d3 = parseInt(tokenStr[2], 10);
      const converted = d1 + d3 / 10;
      if (converted >= 0 && converted <= 10) {
        return Math.round(converted * 100) / 100;
      }
    }

    // 3-digit integer 101..999 on a 10-point grading sheet (e.g., 850 -> 8.5, 725 -> 7.25, 675 -> 6.75)
    if (intVal >= 101 && intVal <= 999) {
      const converted = intVal / 100;
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

  // Replace common handwritten decimal separators between digits (comma, pipe, slash, colon, semicolon) with '.'
  str = str.replace(/(\d)\s*[,|/:;]\s*(\d)/g, "$1.$2").replace(/,/g, ".");

  // Extract all numeric tokens; evaluate from right to left so crossed-out left scores (e.g. "6,8 6,3") yield 6.3
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

export function generateSampleGradingSheetImage(): {
  name: string;
  mimeType: string;
  dataUrl: string;
} {
  return {
    name: "Phieu_Cham_Bai_Thi.png",
    mimeType: "image/png",
    dataUrl: "",
  };
}
