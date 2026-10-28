import { GoogleGenAI, Type } from "@google/genai";

function normalizeSbd(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  let cleaned = String(raw)
    .trim()
    .replace(/\s+/g, "")
    .replace(/[,:;]/g, ".")
    .toUpperCase();

  cleaned = cleaned.replace(/^K[G6O0]\.?(\d+)$/, "K6.$1");
  cleaned = cleaned.replace(/^([A-Z]\d)(\d{3,4})$/, "$1.$2");
  return cleaned;
}

function canonicalSbdKey(raw: unknown): string {
  return normalizeSbd(raw).replace(/[^A-Z0-9]/g, "");
}

function parseVietnameseScore(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number" && !Number.isNaN(raw)) {
    if (raw >= 0 && raw <= 10) return Math.round(raw * 100) / 100;
  }
  const str = String(raw).trim();
  if (!str) return null;

  const normalized = str.replace(/,/g, ".");
  const tokens = normalized.match(/\d+(?:\.\d+)?/g);
  if (!tokens || tokens.length === 0) return null;

  for (let i = tokens.length - 1; i >= 0; i--) {
    const tok = tokens[i];
    const direct = Number.parseFloat(tok);
    if (!Number.isNaN(direct) && direct >= 0 && direct <= 10) {
      return Math.round(direct * 100) / 100;
    }

    if (/^\d{3}$/.test(tok) && tok[1] === "1") {
      const d1 = Number(tok[0]);
      const d3 = Number(tok[2]);
      const candidate = d1 + d3 / 10;
      if (candidate >= 0 && candidate <= 10) {
        return Math.round(candidate * 100) / 100;
      }
    }

    if (/^\d{2}$/.test(tok)) {
      const val = Number(tok);
      if (val > 10 && val <= 99) {
        return Math.round((val / 10) * 100) / 100;
      }
    }
  }

  return null;
}

export default async function handler(req: any, res: any) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const { images, knownSbds } = req.body as {
      images: Array<{ name: string; mimeType: string; data: string }>;
      knownSbds?: string[];
    };

    if (!images || !Array.isArray(images) || images.length === 0) {
      return res.status(400).json({ error: "Vui lòng tải lên ít nhất 1 ảnh phiếu chấm." });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(503).json({
        error: "Chưa cấu hình GEMINI_API_KEY trên máy chủ.",
      });
    }

    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        timeout: 22000,
      },
    });

    const MODEL_POOL = [
      "gemini-3.1-flash-lite",
      "gemini-2.5-flash",
      "gemini-2.0-flash",
    ];

    const knownCanonicalMap = new Map<string, string>();
    if (Array.isArray(knownSbds)) {
      for (const k of knownSbds) {
        const norm = normalizeSbd(k);
        const canon = canonicalSbdKey(k);
        if (canon && norm) knownCanonicalMap.set(canon, norm);
      }
    }

    const extractedRecords: Array<{
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
      const base64Clean = img.data.includes(",") ? img.data.split(",")[1] : img.data;

      let parsedRows: any[] | null = null;
      let lastModelError: any = null;

      for (let attempt = 0; attempt < MODEL_POOL.length; attempt++) {
        const modelName = MODEL_POOL[attempt];
        try {
          const timeoutPromise = new Promise<never>((_, reject) =>
            setTimeout(() => reject(new Error(`Timeout on ${modelName}`)), 16000)
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
Hãy đọc TOÀN BỘ TẤT CẢ CÁC DÒNG từ dòng đầu tiên đến dòng cuối cùng trong bảng của bức ảnh này (bảng thường có từ 20 đến 35 dòng thí sinh, TUYỆT ĐỐI KHÔNG ĐƯỢC DỪNG Ở 5 DÒNG ĐẦU).

QUY TẮC BẮT BUỘC:
1. Trích xuất ĐẦY ĐỦ 100% các dòng có SBD trong bảng (ví dụ từ STT 1 đến STT 24/25/30...). Không bỏ sót bất kỳ dòng nào!
2. Cột "SBD" (Số báo danh): Đọc chính xác chuỗi ký tự in đậm (ví dụ: "K6.001", "K6.025", "K6.048"...).
3. Cột "ĐIỂM BẰNG SỐ" (thường viết tay bằng mực đỏ hoặc xanh):
   - Chú ý đặc trưng chữ viết tay giáo viên Việt Nam:
     + Dấu phẩy "," viết tay nhanh bằng bút đỏ rất hay kéo dài thành nét sổ dọc giống chữ số "1" hoặc "|". Ví dụ: điểm "8,3" nhìn giống "813" hoặc "8|3"; điểm "7,0" nhìn giống "710"; điểm "8,0" nhìn giống "810"; điểm "7,8" nhìn giống "718"; điểm "8,8" nhìn giống "818"; điểm "8,5" nhìn giống "815"; điểm "6,3" nhìn giống "613". Hãy hiểu đó là số thập phân trong thang điểm 0-10 (8.3, 7.0, 8.0, 7.8, 8.8, 8.5, 6.3).
     + Đôi khi giáo viên viết nhanh quên dấu phẩy: "65" nghĩa là 6.5, "83" nghĩa là 8.3.
     + Nếu trong ô có điểm bị gạch xóa ở bên trái và điểm viết lại ở bên phải (ví dụ: "6,8" bị gạch xóa và bên cạnh viết "6,3" hoặc "613"), BẮT BUỘC lấy điểm viết lại hợp lệ ở bên phải (6.3).
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
                            description: "SBD nguyên bản đọc được từ ảnh (VD: K6.025)",
                          },
                          rawScoreText: {
                            type: Type.STRING,
                            description: "Chuỗi điểm viết tay nguyên bản đọc được (VD: 8,5)",
                          },
                          normalizedScore: {
                            type: Type.NUMBER,
                            description: "Điểm số thực từ 0.0 đến 10.0 sau khi quy đổi dấu phẩy thành dấu chấm (VD: 8.5)",
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
        } catch (err: any) {
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
        let sbdNorm = normalizeSbd(item.rawSbd);
        if (!sbdNorm) continue;

        const canon = canonicalSbdKey(sbdNorm);
        if (canon && knownCanonicalMap.has(canon)) {
          sbdNorm = knownCanonicalMap.get(canon)!;
        }

        let finalScore: number | null = null;
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
          id: `${sbdNorm}-${Math.random().toString(36).slice(2, 7)}`,
          sbd: sbdNorm,
          rawSbd: item.rawSbd || sbdNorm,
          score: finalScore,
          rawScoreText: String(item.rawScoreText || finalScore),
          scoreInWords: item.scoreInWords || "",
          confidence:
            item.confidence === "CAO" ||
            item.confidence === "TRUNG_BINH" ||
            item.confidence === "CAN_KIEM_TRA"
              ? item.confidence
              : "CAO",
          notes: item.notes || "Đã chuẩn hóa dấu ',' -> '.'",
          sourceImage: img.name,
        });
      }
    }

    return res.status(200).json({ records: extractedRecords });
  } catch (error: any) {
    return res.status(500).json({
      error: error?.message || "Lỗi trong quá trình nhận diện chữ viết tay OCR.",
    });
  }
}
