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
            score: Math.round((d1 + d2 / 10) * 0.1 * 1000) / 100,
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
 * Extracts and cleans the SBD column (x: 9.5%..23.5%, y: 17.5%..75.5%) and Score column (x: 23.5%..38.5%),
 * removing horizontal table grid lines and upscaling 3x so Tesseract can read all 24-25 rows cleanly.
 */
async function prepareTableColumnCrops(dataUrl: string): Promise<{
  sbdCropUrl: string;
  scoreCropUrl: string;
  fullCleanUrl: string;
}> {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      const W = img.width;
      const H = img.height;
      const srcCanvas = document.createElement("canvas");
      srcCanvas.width = W;
      srcCanvas.height = H;
      const srcCtx = srcCanvas.getContext("2d");
      if (!srcCtx) {
        resolve({ sbdCropUrl: dataUrl, scoreCropUrl: dataUrl, fullCleanUrl: dataUrl });
        return;
      }
      srcCtx.fillStyle = "#FFFFFF";
      srcCtx.fillRect(0, 0, W, H);
      srcCtx.drawImage(img, 0, 0, W, H);
      const rgba = srcCtx.getImageData(0, 0, W, H).data;

      const lum = new Float32Array(W * H);
      for (let i = 0; i < W * H; i++) {
        lum[i] =
          0.299 * rgba[i * 4] +
          0.587 * rgba[i * 4 + 1] +
          0.114 * rgba[i * 4 + 2];
      }

      // 1. SBD Column Crop (x: 10%..23%, y: 17.5%..75%) with horizontal grid-line removal + 3x upscale
      const sX = Math.round(W * 0.098);
      const sY = Math.round(H * 0.175);
      const sW = Math.max(40, Math.round(W * 0.132));
      const sH = Math.max(100, Math.round(H * 0.575));

      const sbdBin = new Uint8Array(sW * sH);
      for (let y = 0; y < sH; y++) {
        const gy = Math.min(H - 1, sY + y);
        let sum = 0;
        for (let x = 0; x < sW; x++) {
          sum += lum[gy * W + Math.min(W - 1, sX + x)];
        }
        const rowMean = sum / sW;
        for (let x = 0; x < sW; x++) {
          const gx = Math.min(W - 1, sX + x);
          const gyUp = Math.max(0, gy - 8);
          const gyDn = Math.min(H - 1, gy + 8);
          const localBg = Math.max(
            rowMean,
            0.5 * (lum[gyUp * W + gx] + lum[gyDn * W + gx])
          );
          sbdBin[y * sW + x] = localBg - lum[gy * W + gx] > 21 ? 0 : 255;
        }
      }

      const sbdCleaned = new Uint8Array(sbdBin);
      for (let y = 3; y < sH - 3; y++) {
        let blackCnt = 0;
        for (let x = 0; x < sW; x++) {
          if (sbdBin[y * sW + x] === 0) blackCnt++;
        }
        if (blackCnt > sW * 0.4) {
          for (let dy = -1; dy <= 1; dy++) {
            const yy = y + dy;
            for (let x = 0; x < sW; x++) {
              if (
                sbdBin[(y - 3) * sW + x] === 0 &&
                sbdBin[(y + 3) * sW + x] === 0
              ) {
                sbdCleaned[yy * sW + x] = 0;
              } else {
                sbdCleaned[yy * sW + x] = 255;
              }
            }
          }
        }
      }

      const scale = 2;
      const sbdCanvas = document.createElement("canvas");
      sbdCanvas.width = sW * scale;
      sbdCanvas.height = sH * scale;
      const sbdCtx = sbdCanvas.getContext("2d");
      let sbdCropUrl = dataUrl;
      if (sbdCtx) {
        const outImg = sbdCtx.createImageData(sW * scale, sH * scale);
        for (let y = 0; y < sH * scale; y++) {
          const sy = Math.floor(y / scale);
          for (let x = 0; x < sW * scale; x++) {
            const sx = Math.floor(x / scale);
            const v = sbdCleaned[sy * sW + sx];
            const idx = (y * sW * scale + x) * 4;
            outImg.data[idx] = v;
            outImg.data[idx + 1] = v;
            outImg.data[idx + 2] = v;
            outImg.data[idx + 3] = 255;
          }
        }
        sbdCtx.putImageData(outImg, 0, 0);
        sbdCropUrl = sbdCanvas.toDataURL("image/png");
      }

      // 2. Score Column Crop (x: 23.2%..37.5%, y: 17.5%..75%) isolating Red/Blue ink
      const cX = Math.round(W * 0.232);
      const cY = Math.round(H * 0.175);
      const cW = Math.max(40, Math.round(W * 0.14));
      const cH = sH;
      const scoreMask = new Uint8Array(cW * cH);

      for (let y = 0; y < cH; y++) {
        const gy = Math.min(H - 1, cY + y);
        let pR = 0,
          pG = 0,
          pB = 0,
          cnt = 0;
        for (let x = Math.max(0, cW - 25); x < cW - 2; x++) {
          const idx = (gy * W + Math.min(W - 1, cX + x)) * 4;
          pR += rgba[idx];
          pG += rgba[idx + 1];
          pB += rgba[idx + 2];
          cnt++;
        }
        if (cnt > 0) {
          pR /= cnt;
          pG /= cnt;
          pB /= cnt;
        }
        const baseRG = pR - pG;
        const baseRB = pR - pB;

        for (let x = 0; x < cW; x++) {
          const idx = (gy * W + Math.min(W - 1, cX + x)) * 4;
          const r = rgba[idx],
            g = rgba[idx + 1],
            b = rgba[idx + 2];
          const darkG = pG - g;
          const redRG = r - g - baseRG;
          const redRB = r - b - baseRB;
          if (darkG > 7 && (redRG > 4 || redRB > 6)) {
            scoreMask[y * cW + x] = 1;
          }
        }
      }

      const scoreCanvas = document.createElement("canvas");
      scoreCanvas.width = cW * scale;
      scoreCanvas.height = cH * scale;
      const scoreCtx = scoreCanvas.getContext("2d");
      let scoreCropUrl = dataUrl;
      if (scoreCtx) {
        const outScore = scoreCtx.createImageData(cW * scale, cH * scale);
        outScore.data.fill(255);
        for (let y = 1; y < cH - 1; y++) {
          for (let x = 1; x < cW - 1; x++) {
            if (scoreMask[y * cW + x]) {
              for (let dy = 0; dy < scale; dy++) {
                for (let dx = 0; dx < scale; dx++) {
                  const uy = y * scale + dy;
                  const ux = x * scale + dx;
                  const idx = (uy * cW * scale + ux) * 4;
                  outScore.data[idx] = 0;
                  outScore.data[idx + 1] = 0;
                  outScore.data[idx + 2] = 0;
                  outScore.data[idx + 3] = 255;
                }
              }
            }
          }
        }
        scoreCtx.putImageData(outScore, 0, 0);
        scoreCropUrl = scoreCanvas.toDataURL("image/png");
      }

      resolve({
        sbdCropUrl,
        scoreCropUrl,
        fullCleanUrl: dataUrl,
      });
    };
    img.onerror = () =>
      resolve({ sbdCropUrl: dataUrl, scoreCropUrl: dataUrl, fullCleanUrl: dataUrl });
    img.src = dataUrl;
  });
}

/**
 * Matches a noisy OCR SBD token (e.g. "K6.025", "K626", "6.046", "K048") to the canonical knownSbds list.
 */
function matchSbdTokenToKnownIndex(
  rawToken: string,
  knownSbds: string[]
): number {
  const cleaned = rawToken
    .toUpperCase()
    .replace(/O/g, "0")
    .replace(/[IL]/g, "1")
    .replace(/S/g, "5")
    .replace(/[^A-Z0-9]/g, "");

  if (!cleaned) return -1;

  // 1. Exact canonical match (e.g. "K6025")
  for (let i = 0; i < knownSbds.length; i++) {
    if (canonicalSbdKey(knownSbds[i]) === cleaned) {
      return i;
    }
  }

  // 2. Match by trailing 3-digit number (e.g., ".025", "6025", "K048" -> "048")
  const trailing3 = cleaned.match(/(\d{3})$/);
  if (trailing3) {
    const suffix = trailing3[1];
    for (let i = 0; i < knownSbds.length; i++) {
      if (canonicalSbdKey(knownSbds[i]).endsWith(suffix)) {
        return i;
      }
    }
  }

  // 3. Match if 2-digit suffix like "K626" -> "026"
  const k2Match = cleaned.match(/^K6(\d{2})$/);
  if (k2Match) {
    const suffix3 = `0${k2Match[1]}`;
    for (let i = 0; i < knownSbds.length; i++) {
      if (canonicalSbdKey(knownSbds[i]).endsWith(suffix3)) {
        return i;
      }
    }
  }

  return -1;
}

/**
 * Runs Local Browser OCR using Column-Cropped Grid-Line Removal + Known-SBD Sequence Alignment.
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
      `Đang tách cột SBD & Điểm trên bảng (${imgIdx + 1}/${images.length}): ${img.name}...`,
      20
    );

    const { sbdCropUrl, scoreCropUrl } = await prepareTableColumnCrops(
      img.dataUrl
    );

    const worker = await Tesseract.createWorker("eng");
    await worker.setParameters({
      tessedit_pageseg_mode: "6" as any,
      tessedit_char_whitelist: "K0123456789.,",
    });

    onProgress?.(
      `Đang đọc cột SBD trên phiếu (${imgIdx + 1}/${images.length}): ${img.name}...`,
      50
    );
    const sbdRes = await worker.recognize(sbdCropUrl);

    await worker.setParameters({
      tessedit_pageseg_mode: "6" as any,
      tessedit_char_whitelist: "0123456789.,",
    });
    onProgress?.(
      `Đang đọc cột Điểm trên phiếu (${imgIdx + 1}/${images.length}): ${img.name}...`,
      80
    );
    const scoreRes = await worker.recognize(scoreCropUrl);
    await worker.terminate();

    const rawSbdLines = (sbdRes.data.text || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l.length >= 2);

    const rawScoreLines = (scoreRes.data.text || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

    // Align detected SBD lines to a contiguous block in knownSbds if available
    let alignedSbds: string[] = [];
    if (knownSbds.length > 0 && rawSbdLines.length >= 3) {
      const offsets: number[] = [];
      rawSbdLines.forEach((line, rowIdx) => {
        const matchedIdx = matchSbdTokenToKnownIndex(line, knownSbds);
        if (matchedIdx >= 0) {
          offsets.push(matchedIdx - rowIdx);
        }
      });

      if (offsets.length >= 2) {
        offsets.sort((a, b) => a - b);
        const startOffset = Math.max(
          0,
          offsets[Math.floor(offsets.length / 2)]
        );
        const rowCount = Math.min(
          Math.max(rawSbdLines.length, 24),
          25,
          knownSbds.length - startOffset
        );
        for (let r = 0; r < rowCount; r++) {
          alignedSbds.push(normalizeSbdClient(knownSbds[startOffset + r]));
        }
      }
    }

    if (alignedSbds.length === 0) {
      for (const line of rawSbdLines) {
        if (knownSbds.length > 0) {
          const idx = matchSbdTokenToKnownIndex(line, knownSbds);
          if (idx >= 0) {
            alignedSbds.push(normalizeSbdClient(knownSbds[idx]));
          }
        } else {
          const norm = normalizeSbdClient(line);
          if (/^[A-Z]\d\.\d{3,4}$/.test(norm)) {
            alignedSbds.push(norm);
          }
        }
      }
    }

    const parsedScores: number[] = [];
    for (const sLine of rawScoreLines) {
      const p = parseVietnameseScoreClient(sLine);
      if (p !== null) {
        parsedScores.push(p);
      }
    }

    for (let r = 0; r < alignedSbds.length; r++) {
      const sbd = alignedSbds[r];
      const scoreVal = parsedScores[r] ?? null;
      if (scoreVal !== null) {
        const rawText = String(scoreVal).replace(".", ",");
        results.push({
          id: `local-${img.name}-${r}-${Date.now()}`,
          sbd,
          rawSbd: sbd,
          score: scoreVal,
          rawScoreText: rawText,
          scoreInWords: "",
          confidence: "TRUNG_BINH",
          notes: `Trích xuất Local OCR (${rawText} → ${scoreVal})`,
          sourceImage: img.name,
        });
      }
    }
  }

  onProgress?.("Hoàn tất trích xuất!", 100);
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
