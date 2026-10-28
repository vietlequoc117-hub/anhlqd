import Tesseract from "tesseract.js";
import {
  OcrRecord,
  normalizeSbdClient,
  canonicalSbdKey,
  parseVietnameseScoreClient,
} from "./sampleGradingSheet";

const VN_WORD_TO_DIGIT: Record<string, number> = {
  không: 0,
  khong: 0,
  một: 1,
  mot: 1,
  mốt: 1,
  hai: 2,
  ba: 3,
  bốn: 4,
  bon: 4,
  tư: 4,
  tu: 4,
  năm: 5,
  nam: 5,
  lăm: 5,
  lam: 5,
  sáu: 6,
  sau: 6,
  bảy: 7,
  bay: 7,
  bẩy: 7,
  tám: 8,
  tam: 8,
  chín: 9,
  chin: 9,
  mười: 10,
  muoi: 10,
};

/**
 * Parses Vietnamese score in words (e.g., "Tám phẩy năm" -> 8.5, "Sáu phẩy bảy" -> 6.7, "Mười điểm" -> 10)
 */
export function parseScoreFromVietnameseWords(text: string): {
  score: number;
  matchedPhrase: string;
} | null {
  const lower = text
    .toLowerCase()
    .replace(/[.,;|()[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!lower) return null;

  const tokens = lower.split(" ");
  for (let i = 0; i < tokens.length; i++) {
    const w1 = tokens[i];
    if (Object.prototype.hasOwnProperty.call(VN_WORD_TO_DIGIT, w1)) {
      const d1 = VN_WORD_TO_DIGIT[w1];
      const next1 = tokens[i + 1] || "";
      const next2 = tokens[i + 2] || "";

      if (
        (next1 === "phẩy" ||
          next1 === "phay" ||
          next1 === "chấm" ||
          next1 === "cham" ||
          next1 === "điểm" ||
          next1 === "diem") &&
        Object.prototype.hasOwnProperty.call(VN_WORD_TO_DIGIT, next2)
      ) {
        const d2 = VN_WORD_TO_DIGIT[next2];
        if (d1 <= 9 && d2 <= 9) {
          return {
            score: Math.round((d1 + d2 / 10) * 100) / 100,
            matchedPhrase: `${w1} ${next1} ${next2}`,
          };
        }
      }

      if (next1 === "rưỡi" || next1 === "ruoi") {
        if (d1 <= 9) {
          return {
            score: d1 + 0.5,
            matchedPhrase: `${w1} ${next1}`,
          };
        }
      }

      if (
        next1 === "điểm" ||
        next1 === "diem" ||
        next1 === "chẵn" ||
        next1 === "chan"
      ) {
        return {
          score: d1,
          matchedPhrase: `${w1} ${next1}`,
        };
      }
    }
  }

  return null;
}

/**
 * Adaptive image preprocessor for smartphone photos of grading sheets (e.g. p3_van.jpg):
 * - Computes local background illumination so gray/shadowed paper becomes pure white (255)
 * - Boosts red/crimson ballpoint pen ink (R >> G, B) and blue/black ink into high-contrast black strokes (0)
 */
async function preprocessGradingSheetForOcr(dataUrl: string): Promise<string> {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      const canvas = document.createElement("canvas");
      const scale = img.width < 1200 ? 1.5 : 1;
      const w = Math.round(img.width * scale);
      const h = Math.round(img.height * scale);
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        resolve(dataUrl);
        return;
      }

      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);

      const imageData = ctx.getImageData(0, 0, w, h);
      const data = imageData.data;

      // Estimate median/mean paper luminance across the image
      let lumSum = 0;
      let sampleCount = 0;
      for (let i = 0; i < data.length; i += 64) {
        lumSum += 0.25 * data[i] + 0.45 * data[i + 1] + 0.3 * data[i + 2];
        sampleCount++;
      }
      const avgPaperLum = sampleCount > 0 ? lumSum / sampleCount : 190;
      const darkThreshold = Math.min(185, Math.max(115, avgPaperLum - 26));

      for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];

        const redness = r - Math.max(g, b);
        const blueness = b - Math.max(r, g);
        const lum = 0.25 * r + 0.45 * g + 0.3 * b;

        let val = 255;
        if (redness > 15 && g < avgPaperLum - 8) {
          // Red ink stroke
          val = Math.max(0, 110 - redness * 3);
        } else if (blueness > 15 && r < avgPaperLum - 8) {
          // Blue ink stroke
          val = Math.max(0, 110 - blueness * 3);
        } else if (lum < darkThreshold) {
          // Printed black SBD / text
          val = Math.max(0, (lum - (darkThreshold - 65)) * 2);
        }

        data[i] = val;
        data[i + 1] = val;
        data[i + 2] = val;
      }

      ctx.putImageData(imageData, 0, 0);
      resolve(canvas.toDataURL("image/png"));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

/**
 * Matches a line of OCR text against known SBDs from the uploaded Excel file (or general SBD pattern)
 */
function findSbdInLine(
  lineText: string,
  knownSbds: string[]
): { sbd: string; rawSbd: string } | null {
  const upperLine = lineText.toUpperCase();

  // 1. General SBD pattern matching first (e.g., K6.025, K6.001, A1.012, etc.)
  const generalMatches = Array.from(
    upperLine.matchAll(/\b([A-Z]{1,3}[0-9OILS]{1,2}\s*[.,_-]\s*[0-9OILS]{2,5})\b/g)
  );
  if (generalMatches.length > 0) {
    const rawCandidate = generalMatches[0][1];
    const normalizedCandidate = normalizeSbdClient(rawCandidate);
    if (knownSbds.length > 0) {
      const targetKey = canonicalSbdKey(normalizedCandidate);
      const exactExcelMatch = knownSbds.find(
        (k) => canonicalSbdKey(k) === targetKey
      );
      if (exactExcelMatch) {
        return {
          sbd: normalizeSbdClient(exactExcelMatch),
          rawSbd: rawCandidate,
        };
      }
    }
    return { sbd: normalizedCandidate, rawSbd: rawCandidate };
  }

  // 2. Try matching against known SBDs from the Step 1 Excel file
  const normalizedTokenStream = upperLine
    .replace(/O/g, "0")
    .replace(/[IL]/g, "1")
    .replace(/,/g, ".");

  if (knownSbds.length > 0) {
    for (const known of knownSbds) {
      const cleanKnown = normalizeSbdClient(known);
      if (!cleanKnown) continue;

      if (
        upperLine.includes(cleanKnown) ||
        normalizedTokenStream.includes(cleanKnown)
      ) {
        return { sbd: cleanKnown, rawSbd: cleanKnown };
      }

      const noDotKnown = canonicalSbdKey(cleanKnown);
      const noDotLine = normalizedTokenStream
        .replace(/\./g, "")
        .replace(/\s+/g, "");
      if (noDotKnown.length >= 4 && noDotLine.includes(noDotKnown)) {
        return { sbd: cleanKnown, rawSbd: cleanKnown };
      }
    }
  }

  // 3. Fallback SBD regex without dots (e.g. K6025 or 6-8 digit SBD)
  const fallbackMatch = normalizedTokenStream.match(
    /\b([A-Z]{1,3}\d{3,6}|\d{5,8})\b/
  );
  if (fallbackMatch) {
    return {
      sbd: normalizeSbdClient(fallbackMatch[1]),
      rawSbd: fallbackMatch[1],
    };
  }

  return null;
}

/**
 * Extracts a Vietnamese exam score (0.0 to 10.0) from a line of text after removing the SBD and leading STT
 */
function extractScoreFromLine(
  lineText: string,
  matchedSbd: string,
  rawSbd: string
): {
  score: number;
  rawScoreText: string;
  scoreInWords: string;
  notes: string;
} | null {
  let remainder = lineText;

  // Remove the raw SBD and normalized SBD from the line
  if (rawSbd) {
    remainder = remainder.replace(rawSbd, " ");
  }
  const sbdPrefixRemoved = matchedSbd.replace(".", "[.,_\\-\\s]*");
  remainder = remainder.replace(new RegExp(sbdPrefixRemoved, "gi"), " ");

  // Remove leading STT integer at start of row
  remainder = remainder.replace(/^\s*\|?\s*\d{1,3}\s*\|?\s+/, " ");

  const wordScore = parseScoreFromVietnameseWords(remainder);

  // Look for explicit decimal or vertical-comma score tokens (e.g., "8,5", "6,8", "813", "710", "65", "83")
  const parsed = parseVietnameseScoreClient(remainder);
  if (parsed !== null) {
    const formattedRaw = String(parsed).replace(".", ",");
    if (wordScore && Math.abs(wordScore.score - parsed) > 0.05) {
      return {
        score: wordScore.score,
        rawScoreText: formattedRaw,
        scoreInWords: wordScore.matchedPhrase,
        notes: `Đối chiếu điểm chữ ("${wordScore.matchedPhrase}") → ${wordScore.score}`,
      };
    }
    return {
      score: parsed,
      rawScoreText: formattedRaw,
      scoreInWords: wordScore?.matchedPhrase || "",
      notes: `Chuẩn hóa điểm (${formattedRaw} → ${parsed})`,
    };
  }

  if (wordScore) {
    const rawFormatted = String(wordScore.score).replace(".", ",");
    return {
      score: wordScore.score,
      rawScoreText: rawFormatted,
      scoreInWords: wordScore.matchedPhrase,
      notes: `Xác thực từ cột Điểm bằng chữ ("${wordScore.matchedPhrase}")`,
    };
  }

  return null;
}

/**
 * Runs Local Browser OCR using Tesseract.js across all rows of each grading sheet image.
 */
export async function runLocalBrowserOcr(params: {
  images: Array<{ name: string; mimeType: string; dataUrl: string }>;
  knownSbds: string[];
  onProgress?: (statusText: string, percent: number) => void;
}): Promise<OcrRecord[]> {
  const { images, knownSbds, onProgress } = params;
  const results: OcrRecord[] = [];

  for (let imgIdx = 0; imgIdx < images.length; imgIdx++) {
    const img = images[imgIdx];
    onProgress?.(
      `Đang tiền xử lý ảnh & cân bằng độ tương phản (${imgIdx + 1}/${images.length}): ${img.name}...`,
      20
    );

    const processedDataUrl = await preprocessGradingSheetForOcr(img.dataUrl);

    const recognition = await Tesseract.recognize(processedDataUrl, "vie+eng", {
      logger: (m) => {
        if (
          m.status === "recognizing text" &&
          typeof m.progress === "number"
        ) {
          const pct = Math.round(25 + m.progress * 70);
          onProgress?.(
            `Đang quét OCR tất cả các dòng trên ${img.name}: ${pct}%`,
            pct
          );
        }
      },
    });

    const rawText = recognition.data.text || "";
    const lines = rawText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l.length >= 2);

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (
        /PHIẾU CHẤM|SỐ BÁO DANH|ĐIỂM BẰNG SỐ|GIÁM KHẢO|CỘNG HÒA|ĐỘC LẬP|CHỮ KÝ/i.test(
          line
        )
      ) {
        continue;
      }

      const foundSbd = findSbdInLine(line, knownSbds);
      if (!foundSbd) continue;

      const extractedScore = extractScoreFromLine(
        line,
        foundSbd.sbd,
        foundSbd.rawSbd
      );
      if (!extractedScore) continue;

      results.push({
        id: `local-${img.name}-${i}-${Date.now()}`,
        sbd: foundSbd.sbd,
        rawSbd: foundSbd.rawSbd,
        score: extractedScore.score,
        rawScoreText: extractedScore.rawScoreText,
        scoreInWords: extractedScore.scoreInWords,
        confidence: "CAO",
        notes: extractedScore.notes,
        sourceImage: img.name,
      });
    }
  }

  onProgress?.("Hoàn tất trích xuất & chuẩn hóa điểm!", 100);
  return results;
}

/**
 * Parses bulk pasted text or table rows (e.g. "K6.001 8,5" or a list of scores corresponding to Excel SBDs)
 */
export function parseBulkPastedScores(
  rawPaste: string,
  knownSbds: string[]
): OcrRecord[] {
  const lines = rawPaste
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const records: OcrRecord[] = [];

  lines.forEach((line, idx) => {
    const parts = line.split(/[\t;|-]+|\s{2,}|\s+/);
    if (parts.length >= 2) {
      const possibleSbd = normalizeSbdClient(parts[0]);
      const rawScoreStr = parts[parts.length - 1];
      const parsedScore = parseVietnameseScoreClient(rawScoreStr);

      if (
        possibleSbd &&
        parsedScore !== null &&
        /[A-Z0-9]{3,}/.test(possibleSbd)
      ) {
        records.push({
          id: `paste-${idx}-${Date.now()}`,
          sbd: possibleSbd,
          rawSbd: parts[0],
          score: parsedScore,
          rawScoreText: rawScoreStr,
          scoreInWords: "",
          confidence: "CAO",
          notes: `Dán nhanh: "${rawScoreStr}" → ${parsedScore}`,
          sourceImage: "Dán hàng loạt",
        });
        return;
      }
    }

    const singleScore = parseVietnameseScoreClient(line);
    if (singleScore !== null && idx < knownSbds.length) {
      const targetSbd = normalizeSbdClient(knownSbds[idx]);
      records.push({
        id: `paste-seq-${idx}-${Date.now()}`,
        sbd: targetSbd,
        rawSbd: targetSbd,
        score: singleScore,
        rawScoreText: line,
        scoreInWords: "",
        confidence: "CAO",
        notes: `Gán theo thứ tự SBD #${idx + 1}: "${line}" → ${singleScore}`,
        sourceImage: "Dán theo thứ tự SBD",
      });
    }
  });

  return records;
}
