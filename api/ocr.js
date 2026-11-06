import { GoogleGenAI, Type } from "@google/genai";

function normalizeSbd(raw) {
  if (raw === null || raw === undefined) return "";
  let s = String(raw)
    .trim()
    .replace(/\s+/g, "")
    .replace(/[,_-]+/g, ".")
    .toUpperCase();

  s = s.replace(/^K[G6O0]\.?(\d+)$/, "K6.$1");
  s = s.replace(/^([A-Z]\d)(\d{3,4})$/, "$1.$2");
  s = s.replace(/\.([0-9OILS]{2,6})$/, (_, suffix) => {
    const fixed = suffix
      .replace(/O/g, "0")
      .replace(/[IL]/g, "1")
      .replace(/S/g, "5");
    return `.${fixed}`;
  });
  s = s.replace(/^([A-Z]{1,3})([0-9OILS]{1,5})$/, (_, prefix, digits) => {
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

function canonicalSbdKey(raw) {
  return normalizeSbd(raw).replace(/[^A-Z0-9]/g, "");
}

function parseVietnameseScore(raw) {
  if (raw === null || raw === undefined) return null;

  const normalizeSingleNumericValue = (val, rawToken) => {
    if (Number.isNaN(val) || val < 0) return null;
    if (val <= 10) {
      return Math.round(val * 100) / 100;
    }
    const intVal = Math.round(val);
    const tokenStr = rawToken ? rawToken.replace(/[^0-9]/g, "") : String(intVal);

    if (intVal === 100 || intVal === 1000) return 10;

    // 4-digit number where the 2nd digit is '1' (tall comma stroke in 2-decimal score, e.g., 8150 -> 8.5)
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

    // 3-digit integer 101..999 on a 10-point grading sheet (e.g., 850 -> 8.5, 725 -> 7.25)
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

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "15mb",
    },
  },
  maxDuration: 60,
};

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const apiKey =
    process.env.GEMINI_API_KEY ||
    process.env.VITE_GEMINI_API_KEY ||
    process.env.GOOGLE_API_KEY ||
    process.env.GOOGLE_GENERATIVE_AI_API_KEY;

  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      runtime: "vercel-serverless",
      hasApiKey: Boolean(apiKey),
    });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
    const { images, knownSbds, phachToSbdMap } = body;

    if (!images || !Array.isArray(images) || images.length === 0) {
      return res.status(400).json({ error: "Vui lòng tải lên ít nhất 1 ảnh phiếu chấm." });
    }

    if (!apiKey) {
      return res.status(503).json({
        errorCode: "MISSING_GEMINI_API_KEY",
        error:
          "Chưa cấu hình GEMINI_API_KEY trên Vercel (Project Settings -> Environment Variables -> GEMINI_API_KEY).",
      });
    }

    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        timeout: 24000,
      },
    });

    const MODEL_POOL = [
      "gemini-3.1-flash-lite",
      "gemini-2.5-flash",
      "gemini-3.1-flash-lite",
    ];

    const knownCanonicalMap = new Map();
    if (Array.isArray(knownSbds)) {
      for (const k of knownSbds) {
        const norm = normalizeSbd(k);
        const canon = canonicalSbdKey(k);
        if (canon && norm) knownCanonicalMap.set(canon, norm);
      }
    }

    const phachMap = new Map();
    if (phachToSbdMap && typeof phachToSbdMap === "object") {
      for (const [pKey, sVal] of Object.entries(phachToSbdMap)) {
        const normP = normalizeSbd(pKey);
        const canonP = canonicalSbdKey(pKey);
        const normS = normalizeSbd(sVal);
        if (normP && normS) phachMap.set(normP, normS);
        if (canonP && normS) phachMap.set(canonP, normS);
      }
    }

    const extractedRecords = [];

    for (const img of images) {
      const base64Clean = img.data.includes(",") ? img.data.split(",")[1] : img.data;

      let parsedRows = null;
      let lastModelError = null;

      for (let attempt = 0; attempt < MODEL_POOL.length; attempt++) {
        const modelName = MODEL_POOL[attempt];
        try {
          const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error(`Timeout on ${modelName}`)), 18000)
          );

          const response = await Promise.race([
            ai.models.generateContent({
              model: modelName,
              contents: {
                parts: [
                  {
                    inlineData: {
                      mimeType: img.mimeType || "image/jpeg",
                      data: base64Clean,
                    },
                  },
                  {
                    text: `Bạn là chuyên gia OCR bảng điểm và "PHIẾU CHẤM BÀI THI" tại Việt Nam.
Hãy đọc TOÀN BỘ TẤT CẢ CÁC DÒNG từ dòng đầu tiên đến dòng cuối cùng trong bảng của bức ảnh này (bảng có thể có 5, 20, 25 hoặc 35+ dòng, TUYỆT ĐỐI KHÔNG ĐƯỢC DỪNG Ở 5 DÒNG ĐẦU).

QUY TẮC BẮT BUỘC:
1. Trích xuất ĐẦY ĐỦ 100% các dòng có ghi điểm trong bảng. Bỏ qua các dòng mà cả cột MÃ PHÁCH/SBD và cột ĐIỂM BẰNG SỐ đều để trống (ví dụ thí sinh vắng thi).
2. Cột mã định danh bài thi (có tiêu đề là "SBD" hoặc "MÃ PHÁCH"):
   - Nếu bảng dùng cột "MÃ PHÁCH" (ví dụ: "A40", "A41", "A42", "A43", "B01"...), trả về chính xác mã phách đó vào trường "rawSbd" và "maPhach".
   - Nếu bảng dùng cột "SBD" (ví dụ: "K6.001", "K6.025", "K9.001"...), trả về chính xác SBD đó vào trường "rawSbd".
   - Nếu bảng có cả 2 cột "SBD" và "MÃ PHÁCH", trả về cả hai tương ứng.
3. Cột "ĐIỂM BẰNG SỐ" (thường viết tay bằng mực đỏ hoặc xanh):
   - Chú ý đặc trưng chữ viết tay giáo viên Việt Nam:
     + Điểm nguyên hoặc có 2 chữ số thập phân như "8", "8,50", "8,5", "7,25": đọc chính xác thành 8, 8.5, 8.5, 7.25.
     + Dấu phẩy "," viết tay nhanh bằng bút đỏ rất hay kéo dài thành nét sổ dọc giống chữ số "1" hoặc "|". Ví dụ: điểm "8,3" nhìn giống "813" hoặc "8|3"; điểm "7,0" nhìn giống "710"; điểm "8,0" nhìn giống "810"; điểm "7,8" nhìn giống "718"; điểm "8,8" nhìn giống "818"; điểm "8,5" nhìn giống "815"; điểm "6,3" nhìn giống "613". Hãy hiểu đó là số thập phân trong thang điểm 0-10 (8.3, 7.0, 8.0, 7.8, 8.8, 8.5, 6.3).
     + Đôi khi giáo viên viết nhanh quên dấu phẩy: "65" nghĩa là 6.5, "83" nghĩa là 8.3, "850" nghĩa là 8.50 (8.5).
     + Nếu trong ô có điểm bị gạch xóa ở bên trái và điểm viết lại ở bên phải (ví dụ: "6,8" bị gạch xóa và bên cạnh viết "6,3"), BẮT BUỘC lấy điểm viết lại hợp lệ ở bên phải (6.3).
4. Quy đổi chuẩn xác sang số thực dấu chấm động (float) nằm trong khoảng từ 0.0 đến 10.0.
5. Trả về danh sách JSON theo thứ tự từ trên xuống dưới.`,
                  },
                ],
              },
              config: {
                responseMimeType: "application/json",
                responseSchema: {
                  type: Type.OBJECT,
                  properties: {
                    rows: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          rawSbd: {
                            type: Type.STRING,
                            description: "Mã SBD hoặc Mã Phách nguyên bản đọc được từ ảnh (VD: A40, A41, K9.001, K6.025)",
                          },
                          maPhach: {
                            type: Type.STRING,
                            description: "Mã phách nếu cột trên phiếu là MÃ PHÁCH (VD: A40, A41)",
                          },
                          rawScoreText: {
                            type: Type.STRING,
                            description: "Chuỗi điểm viết tay nguyên bản đọc được (VD: 8, 8,50, 8,5)",
                          },
                          normalizedScore: {
                            type: Type.NUMBER,
                            description: "Điểm số thực từ 0.0 đến 10.0 sau khi quy đổi dấu phẩy thành dấu chấm (VD: 8, 8.5)",
                          },
                          scoreInWords: {
                            type: Type.STRING,
                            description: "Cột điểm bằng chữ nếu có",
                          },
                          confidence: {
                            type: Type.STRING,
                            description: "CAO, TRUNG_BINH, hoặc CAN_KIEM_TRA",
                          },
                          notes: {
                            type: Type.STRING,
                            description: "Ghi chú nếu có gạch xóa hoặc sửa điểm",
                          },
                        },
                        required: ["rawSbd", "rawScoreText", "normalizedScore", "confidence"],
                      },
                    },
                  },
                  required: ["rows"],
                },
              },
            }),
            timeoutPromise,
          ]);

          const jsonText = response.text || "{}";
          const parsed = JSON.parse(jsonText);
          if (parsed && Array.isArray(parsed.rows) && parsed.rows.length > 0) {
            parsedRows = parsed.rows;
            break;
          }
        } catch (err) {
          lastModelError = err;
          if (attempt < MODEL_POOL.length - 1) {
            await new Promise((r) => setTimeout(r, 400));
          }
        }
      }

      if (!parsedRows) {
        throw lastModelError || new Error("Không thể nhận diện ảnh phiếu chấm.");
      }

      for (const item of parsedRows) {
        let codeNorm = normalizeSbd(item.rawSbd || item.maPhach);
        if (!codeNorm) continue;

        const canon = canonicalSbdKey(codeNorm);
        if (canon && knownCanonicalMap.has(canon)) {
          codeNorm = knownCanonicalMap.get(canon);
        }

        let detectedPhach = normalizeSbd(item.maPhach || "");
        let finalSbd = codeNorm;

        // Check if codeNorm or detectedPhach maps to an SBD via phachToSbdMap (Hình 2)
        if (phachMap.has(codeNorm) || phachMap.has(canon)) {
          detectedPhach = codeNorm;
          finalSbd = phachMap.get(codeNorm) || phachMap.get(canon);
        } else if (detectedPhach && (phachMap.has(detectedPhach) || phachMap.has(canonicalSbdKey(detectedPhach)))) {
          finalSbd =
            phachMap.get(detectedPhach) ||
            phachMap.get(canonicalSbdKey(detectedPhach));
        }

        let finalScore = null;
        if (
          typeof item.normalizedScore === "number" &&
          !Number.isNaN(item.normalizedScore) &&
          item.normalizedScore >= 0 &&
          item.normalizedScore <= 10
        ) {
          finalScore = Math.round(item.normalizedScore * 100) / 100;
        } else {
          finalScore =
            parseVietnameseScore(item.rawScoreText) ??
            parseVietnameseScore(item.normalizedScore);
        }

        if (finalScore === null) continue;

        extractedRecords.push({
          id: `${finalSbd}-${Math.random().toString(36).slice(2, 7)}`,
          sbd: finalSbd,
          rawSbd: item.rawSbd || codeNorm,
          maPhach: detectedPhach || (finalSbd !== codeNorm ? codeNorm : undefined),
          score: finalScore,
          rawScoreText: String(item.rawScoreText || finalScore),
          scoreInWords: item.scoreInWords || "",
          confidence:
            item.confidence === "CAO" ||
            item.confidence === "TRUNG_BINH" ||
            item.confidence === "CAN_KIEM_TRA"
              ? item.confidence
              : "CAO",
          notes:
            detectedPhach && finalSbd !== detectedPhach
              ? `Khớp Mã phách ${detectedPhach} → SBD ${finalSbd} (${item.rawScoreText} → ${finalScore})`
              : item.notes || `Chuẩn hóa (${item.rawScoreText} → ${finalScore})`,
          sourceImage: img.name,
        });
      }
    }

    return res.status(200).json({ records: extractedRecords });
  } catch (error) {
    return res.status(500).json({
      error: error?.message || "Lỗi trong quá trình nhận diện chữ viết tay OCR.",
    });
  }
}
