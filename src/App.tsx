import React, { useState, useMemo, useRef } from "react";
import {
  Upload,
  FileSpreadsheet,
  ScanText,
  Download,
  CheckCircle2,
  AlertCircle,
  Code2,
  Copy,
  Check,
  Plus,
  Trash2,
  RefreshCw,
  Search,
  ZoomIn,
  ZoomOut,
  FileCheck2,
  FileUp,
  ArrowRight,
  ArrowLeft,
  Cpu,
  Keyboard,
  ClipboardPaste,
} from "lucide-react";
import {
  OcrRecord,
  ExcelSheetData,
  normalizeSbdClient,
  parseVietnameseScoreClient,
  generateSampleGradingSheetImage,
} from "./utils/sampleGradingSheet";
import {
  parseExcelClientSide,
  exportUpdatedExcelClientSide,
  generateSampleExcelClientSide,
} from "./utils/clientExcelEngine";
import {
  runLocalBrowserOcr,
  parseBulkPastedScores,
} from "./utils/localOcrEngine";

interface UploadedImage {
  id: string;
  name: string;
  mimeType: string;
  dataUrl: string;
  status?: "PENDING" | "PROCESSING" | "DONE" | "ERROR";
  extractedCount?: number;
}

/**
 * Compresses large smartphone photos client-side (max 1600px) so uploading 13–50 images
 * simultaneously never hits 413 Payload Too Large limits on Cloud Run or Vercel.
 */
async function optimizeImageForBatchOcr(
  file: File
): Promise<{ dataUrl: string; mimeType: string }> {
  const rawDataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const maxDim = 1600;
      if (img.width <= maxDim && img.height <= maxDim && file.size < 900 * 1024) {
        resolve({ dataUrl: rawDataUrl, mimeType: file.type || "image/png" });
        return;
      }
      const scale = Math.min(maxDim / img.width, maxDim / img.height, 1);
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        resolve({ dataUrl: rawDataUrl, mimeType: file.type || "image/png" });
        return;
      }
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve({
        dataUrl: canvas.toDataURL("image/jpeg", 0.88),
        mimeType: "image/jpeg",
      });
    };
    img.onerror = () =>
      resolve({ dataUrl: rawDataUrl, mimeType: file.type || "image/png" });
    img.src = rawDataUrl;
  });
}

type OcrEngineMode = "LOCAL_WASM" | "RAPID_KEYBOARD" | "HYBRID_CLOUD";

export default function App() {
  // Sequential Step Navigation (1: Excel Nguồn -> 2: Quét Ảnh & Điền Điểm Excel)
  const [currentStep, setCurrentStep] = useState<1 | 2>(1);

  // STEP 1: Excel Source File State (100% Client-Side Capable)
  const [excelData, setExcelData] = useState<ExcelSheetData | null>(null);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [headerRow, setHeaderRow] = useState<number>(5);
  const [sbdCol, setSbdCol] = useState<string>("G");
  const [scoreCol, setScoreCol] = useState<string>("I");
  const [isExcelLoading, setIsExcelLoading] = useState<boolean>(false);
  const [excelError, setExcelError] = useState<string | null>(null);
  const [isExcelDragging, setIsExcelDragging] = useState<boolean>(false);

  // STEP 2: Grading Sheet Images & Zero-API / Hybrid OCR State
  const [engineMode, setEngineMode] = useState<OcrEngineMode>("HYBRID_CLOUD");
  const [images, setImages] = useState<UploadedImage[]>([]);
  const [activeImageIdx, setActiveImageIdx] = useState<number>(0);
  const [imageZoom, setImageZoom] = useState<number>(1);
  const [isOcrRunning, setIsOcrRunning] = useState<boolean>(false);
  const [ocrProgressText, setOcrProgressText] = useState<string>("");
  const [ocrProgressPct, setOcrProgressPct] = useState<number>(0);
  const [ocrNotice, setOcrNotice] = useState<string | null>(null);
  const [ocrRecords, setOcrRecords] = useState<OcrRecord[]>([]);

  // Rapid Keyboard Entry & Bulk Paste State (Zero-API Mode)
  const [rapidInputs, setRapidInputs] = useState<Record<string, string>>({});
  const [activeRapidRowIdx, setActiveRapidRowIdx] = useState<number>(0);
  const [showBulkPaste, setShowBulkPaste] = useState<boolean>(false);
  const [bulkPasteText, setBulkPasteText] = useState<string>("");
  const rapidInputRefs = useRef<Record<number, HTMLInputElement | null>>({});

  // STEP 3: Table Filtering & Export State
  const [rowFilter, setRowFilter] = useState<"ALL" | "MATCHED" | "UNMATCHED">("ALL");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [exportResult, setExportResult] = useState<{
    updatedExcelBase64: string;
    outputFileName: string;
    matchedCount: number;
    untouchedCount: number;
    pythonAuditLog: string;
    openpyxlCode: string;
  } | null>(null);
  const [copiedCode, setCopiedCode] = useState<boolean>(false);

  const excelInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);

  // Extract list of known SBDs from Step 1 Excel file to anchor Local OCR
  const knownSbds = useMemo(() => {
    if (!excelData) return [];
    return excelData.studentRows
      .map((r) => normalizeSbdClient(r.cells[sbdCol] ?? r.sbd))
      .filter(Boolean);
  }, [excelData, sbdCol]);

  // Lookup map for fast VLOOKUP matching
  const ocrLookupMap = useMemo(() => {
    const map = new Map<string, OcrRecord>();
    for (const rec of ocrRecords) {
      const norm = normalizeSbdClient(rec.sbd);
      if (norm) map.set(norm, rec);
    }
    return map;
  }, [ocrRecords]);

  // Match Excel student rows with OCR results
  const matchedStudentRows = useMemo(() => {
    if (!excelData) return [];
    return excelData.studentRows.map((row) => {
      const rawSbd = row.cells[sbdCol] ?? row.sbd;
      const normSbd = normalizeSbdClient(rawSbd);
      const matchedOcr = normSbd ? ocrLookupMap.get(normSbd) : undefined;
      return {
        ...row,
        displaySbd: rawSbd !== null && rawSbd !== undefined ? String(rawSbd) : "",
        normalizedSbd: normSbd,
        matchedScore: matchedOcr ? matchedOcr.score : null,
        matchedRawText: matchedOcr ? matchedOcr.rawScoreText : null,
        matchedSource: matchedOcr ? matchedOcr.sourceImage : null,
        isMatched: Boolean(matchedOcr),
      };
    });
  }, [excelData, sbdCol, ocrLookupMap]);

  const filteredStudentRows = useMemo(() => {
    return matchedStudentRows.filter((r) => {
      if (rowFilter === "MATCHED" && !r.isMatched) return false;
      if (rowFilter === "UNMATCHED" && r.isMatched) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.trim().toLowerCase();
        const rowText = Object.values(r.cells)
          .map((v) => String(v ?? ""))
          .join(" ")
          .toLowerCase();
        return rowText.includes(q) || r.displaySbd.toLowerCase().includes(q);
      }
      return true;
    });
  }, [matchedStudentRows, rowFilter, searchQuery]);

  const stats = useMemo(() => {
    const totalExcelStudents = matchedStudentRows.length;
    const matchedCount = matchedStudentRows.filter((r) => r.isMatched).length;
    const unmatchedExcelCount = totalExcelStudents - matchedCount;
    return {
      totalOcrRows: ocrRecords.length,
      totalExcelStudents,
      matchedCount,
      unmatchedExcelCount,
    };
  }, [matchedStudentRows, ocrRecords]);

  // Openpyxl Python code preview
  const liveOpenpyxlScript = useMemo(() => {
    if (exportResult?.openpyxlCode) return exportResult.openpyxlCode;
    const fileName = excelData?.fileName || "Danh_Sach_Hoc_Sinh.xlsx";
    const startRow = headerRow + 1;
    const dictLines = ocrRecords
      .map(
        (r) =>
          `    "${normalizeSbdClient(r.sbd)}": ${r.score},  # Gốc trên phiếu: "${r.rawScoreText}"`
      )
      .join("\n");

    return `import re
import openpyxl

def normalize_sbd(value) -> str:
    """Chuẩn hóa SBD (loại bỏ khoảng trắng thừa, viết hoa)."""
    if value is None:
        return ""
    return re.sub(r"\\s+", "", str(value).strip()).replace(",", ".").upper()

def sync_scores_to_excel(excel_path: str, output_path: str):
    # BƯỚC 1: Mở file Excel nguồn "${fileName}", bảo toàn 100% định dạng gốc
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    sheet_names = wb.sheetnames
    ws = wb[sheet_names[${selectedSheetIndex}]] if len(sheet_names) > ${selectedSheetIndex} else wb.active

    HEADER_ROW = ${headerRow}
    START_DATA_ROW = ${startRow}
    SBD_COL = "${sbdCol}"
    SCORE_COL = "${scoreCol}"

    # BƯỚC 2: Dữ liệu (SBD -> Điểm số Float) đã trích xuất từ PHIẾU CHẤM BÀI THI
    ocr_scores: dict[str, float] = {
${dictLines || "    # Dữ liệu điểm từ Phiếu Chấm Bài Thi"}
    }

    updated_count = 0

    # BƯỚC 3: Quét cột ${sbdCol} từ hàng ${startRow} và ghi điểm Float vào cột ${scoreCol}
    for row_idx in range(START_DATA_ROW, ws.max_row + 1):
        raw_sbd = ws[f"{SBD_COL}{row_idx}"].value
        norm_sbd = normalize_sbd(raw_sbd)

        if norm_sbd and norm_sbd in ocr_scores:
            # Ghi đè duy nhất giá trị điểm dạng số thực (Float)
            ws[f"{SCORE_COL}{row_idx}"].value = float(ocr_scores[norm_sbd])
            updated_count += 1

    wb.save(output_path)
    print(f"Hoàn tất! Đã điền điểm {updated_count} học sinh vào file: {output_path}")

if __name__ == "__main__":
    sync_scores_to_excel("${fileName}", "Da_Nhap_Diem_${fileName}")`;
  }, [exportResult, excelData, selectedSheetIndex, headerRow, sbdCol, scoreCol, ocrRecords]);

  // ============================================================================
  // STEP 1 HANDLERS: 100% Client-Side Excel Parsing (Works on Vercel & Offline)
  // ============================================================================
  const parseAndLoadExcelBuffer = async (
    file: File | Blob,
    fileName: string,
    sheetIndex = 0
  ) => {
    setIsExcelLoading(true);
    setExcelError(null);

    try {
      const base64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });

      const parsed = await parseExcelClientSide({
        fileBase64: base64,
        fileName,
        targetSheetIndex: sheetIndex,
      });

      setExcelData(parsed);
      setSelectedSheetIndex(sheetIndex);
      setHeaderRow(parsed.detectedHeaderRow || 5);
      setSbdCol(parsed.detectedSbdCol || "G");
      setScoreCol(parsed.detectedScoreCol || "I");
      setExportResult(null);

      // Move to Step 2 automatically after loading Source Excel file
      setCurrentStep(2);
    } catch (err: any) {
      setExcelError(
        err?.message ||
          "Không thể đọc file Excel. Vui lòng đảm bảo file có định dạng .xlsx chuẩn."
      );
    } finally {
      setIsExcelLoading(false);
      if (excelInputRef.current) excelInputRef.current.value = "";
    }
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await parseAndLoadExcelBuffer(file, file.name, 0);
  };

  const handleExcelDrop = async (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsExcelDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (!file) return;
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      setExcelError("Vui lòng tải lên file Excel có định dạng .xlsx.");
      return;
    }
    await parseAndLoadExcelBuffer(file, file.name, 0);
  };

  const handleSheetChange = async (newSheetIdx: number) => {
    if (!excelData) return;
    setSelectedSheetIndex(newSheetIdx);
    setIsExcelLoading(true);
    try {
      const parsed = await parseExcelClientSide({
        fileBase64: excelData.fileBase64,
        fileName: excelData.fileName,
        targetSheetIndex: newSheetIdx,
      });
      setExcelData(parsed);
      setHeaderRow(parsed.detectedHeaderRow || 5);
      setSbdCol(parsed.detectedSbdCol || "G");
      setScoreCol(parsed.detectedScoreCol || "I");
    } catch (err: any) {
      setExcelError(err?.message || "Không thể chuyển trang tính.");
    } finally {
      setIsExcelLoading(false);
    }
  };

  const handleHeaderRowChange = async (newHeaderRow: number) => {
    setHeaderRow(newHeaderRow);
    if (!excelData) return;
    try {
      const parsed = await parseExcelClientSide({
        fileBase64: excelData.fileBase64,
        fileName: excelData.fileName,
        preferredHeaderRow: newHeaderRow,
        preferredSbdCol: sbdCol,
        preferredScoreCol: scoreCol,
        targetSheetIndex: selectedSheetIndex,
      });
      setExcelData(parsed);
    } catch {
      // Keep current state
    }
  };

  // Load Sample Bundle (Hình 2 Excel + Hình 1 Grading Sheet -> Matched Hình 3)
  const handleLoadSampleWorkspace = async () => {
    setIsExcelLoading(true);
    setOcrNotice(null);
    setExcelError(null);

    try {
      // 1. Load Hình 2: Trường THCS Mỹ Xá - Bảng Điểm Thi Môn: Văn 6 (K6.001..K6.005)
      const sampleExcel = await generateSampleExcelClientSide();
      setExcelData(sampleExcel);
      setSelectedSheetIndex(0);
      setHeaderRow(sampleExcel.detectedHeaderRow || 5);
      setSbdCol(sampleExcel.detectedSbdCol || "G");
      setScoreCol(sampleExcel.detectedScoreCol || "I");

      // 2. Load Hình 1: Actual uploaded grading sheet photo (/assets/aistudio/image_74b23c14.png)
      let dataUrl = "";
      try {
        const resp = await fetch("/assets/aistudio/image_74b23c14.png");
        if (resp.ok) {
          const blob = await resp.blob();
          dataUrl = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result));
            reader.onerror = reject;
            reader.readAsDataURL(blob);
          });
        }
      } catch {
        // Fallback to generated canvas if static asset is unavailable
      }
      if (!dataUrl) {
        dataUrl = generateSampleGradingSheetImage().dataUrl;
      }

      const sampleImageObj: UploadedImage = {
        id: `hinh1-${Date.now()}`,
        name: "Hinh_1_Phieu_Cham_Van6.png",
        mimeType: "image/png",
        dataUrl,
      };
      setImages([sampleImageObj]);
      setActiveImageIdx(0);

      const sbdsFromSample = sampleExcel.studentRows
        .map((r) => normalizeSbdClient(r.cells["G"] ?? r.sbd))
        .filter(Boolean);

      setCurrentStep(2);
      await executeOcrPipeline([sampleImageObj], sbdsFromSample, false);
    } catch (err: any) {
      setExcelError(err?.message || "Lỗi khi khởi tạo dữ liệu mẫu.");
    } finally {
      setIsExcelLoading(false);
    }
  };

  // ============================================================================
  // STEP 2 HANDLERS: Local WebAssembly OCR / Rapid Keyboard / Cloud Fallback
  // ============================================================================
  const mergeIncomingRecords = (incoming: OcrRecord[], append: boolean) => {
    setOcrRecords((prev) => {
      const combined = append ? [...prev, ...incoming] : incoming;
      const bySbd = new Map<string, OcrRecord>();
      for (const item of combined) {
        const key = normalizeSbdClient(item.sbd);
        if (key) bySbd.set(key, item);
      }
      const nextList = Array.from(bySbd.values());

      // Sync rapid keyboard inputs map as well
      const nextRapid: Record<string, string> = {};
      for (const r of nextList) {
        nextRapid[normalizeSbdClient(r.sbd)] = r.rawScoreText;
      }
      setRapidInputs((prevRapid) => ({ ...prevRapid, ...nextRapid }));
      return nextList;
    });
    setExportResult(null);
  };

  const executeOcrPipeline = async (
    targetImages: UploadedImage[],
    activeKnownSbds: string[],
    append = false
  ) => {
    if (targetImages.length === 0) return;

    if (engineMode === "RAPID_KEYBOARD") {
      setOcrNotice(
        "Đang ở Chế độ Nhập Nhanh Bàn Phím (Không cần API): Gõ điểm (VD: 8,5) và nhấn Enter để xuống dòng tức thì."
      );
      return;
    }

    setIsOcrRunning(true);
    setOcrNotice(null);
    setOcrProgressPct(5);

    if (!append) {
      setOcrRecords([]);
    }

    let totalExtractedAcrossBatch = 0;
    const totalFiles = targetImages.length;

    try {
      // Process each image sequentially in a fault-tolerant batch queue (1..13+)
      // so each sheet streams into the table immediately and never triggers 413/timeout!
      for (let idx = 0; idx < totalFiles; idx++) {
        const currentImg = targetImages[idx];
        const basePct = Math.round((idx / totalFiles) * 100);

        // Mark this image tab as PROCESSING and focus it so the user sees live progress
        setImages((prev) =>
          prev.map((item) =>
            item.id === currentImg.id ? { ...item, status: "PROCESSING" } : item
          )
        );
        setActiveImageIdx((prevIdx) => {
          const foundIdx = images.findIndex((x) => x.id === currentImg.id);
          return foundIdx >= 0 ? foundIdx : prevIdx;
        });

        setOcrProgressText(
          `Đang xử lý phiếu chấm (${idx + 1}/${totalFiles}): ${currentImg.name}...`
        );
        setOcrProgressPct(Math.max(8, basePct));

        let sheetRecords: OcrRecord[] = [];

        if (engineMode === "LOCAL_WASM") {
          sheetRecords = await runLocalBrowserOcr({
            images: [currentImg],
            knownSbds: activeKnownSbds,
            onProgress: (text, subPct) => {
              const overallPct = Math.min(
                99,
                Math.round(((idx + subPct / 100) / totalFiles) * 100)
              );
              setOcrProgressText(`[Phiếu ${idx + 1}/${totalFiles}] ${text}`);
              setOcrProgressPct(overallPct);
            },
          });
        } else {
          // HYBRID_CLOUD: Call /api/ocr per single compressed image with automatic fallback to Local OCR
          try {
            const response = await fetch("/api/ocr", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                images: [
                  {
                    name: currentImg.name,
                    mimeType: currentImg.mimeType,
                    data: currentImg.dataUrl,
                  },
                ],
              }),
            });

            if (!response.ok) {
              throw new Error(`HTTP ${response.status}`);
            }
            const data = await response.json();
            if (Array.isArray(data.records) && data.records.length > 0) {
              sheetRecords = data.records;
            } else {
              throw new Error("Empty cloud records");
            }
          } catch {
            // Automatic per-image fallback to Local Browser OCR if cloud is busy or offline
            sheetRecords = await runLocalBrowserOcr({
              images: [currentImg],
              knownSbds: activeKnownSbds,
              onProgress: (text, subPct) => {
                const overallPct = Math.min(
                  99,
                  Math.round(((idx + subPct / 100) / totalFiles) * 100)
                );
                setOcrProgressText(
                  `[Phiếu ${idx + 1}/${totalFiles} · Dự phòng Local] ${text}`
                );
                setOcrProgressPct(overallPct);
              },
            });
          }
        }

        totalExtractedAcrossBatch += sheetRecords.length;

        // Immediately stream this sheet's extracted scores into the live table & Hình 3 preview!
        if (sheetRecords.length > 0) {
          mergeIncomingRecords(sheetRecords, true);
        }

        // Mark this image tab as DONE with extracted count
        setImages((prev) =>
          prev.map((item) =>
            item.id === currentImg.id
              ? {
                  ...item,
                  status: "DONE",
                  extractedCount: sheetRecords.length,
                }
              : item
          )
        );

        setOcrProgressPct(Math.round(((idx + 1) / totalFiles) * 100));
      }

      setOcrNotice(
        `Đã quét hoàn tất ${totalFiles} phiếu chấm và trích xuất tổng cộng ${totalExtractedAcrossBatch} dòng điểm.`
      );
    } catch (err: any) {
      setOcrNotice(
        err?.message ||
          "Đã xử lý xong các phiếu khả dụng. Bạn có thể kiểm tra kết quả ở bảng bên phải."
      );
    } finally {
      setIsOcrRunning(false);
      setOcrProgressText("");
      setOcrProgressPct(0);
    }
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setOcrNotice(`Đang nén tối ưu & nạp ${files.length} ảnh phiếu chấm...`);
    const newImages: UploadedImage[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const optimized = await optimizeImageForBatchOcr(file);
      newImages.push({
        id: `${file.name}-${Date.now()}-${i}`,
        name: file.name,
        mimeType: optimized.mimeType,
        dataUrl: optimized.dataUrl,
        status: "PENDING",
      });
    }

    const updatedList = [...images, ...newImages];
    setImages(updatedList);
    setActiveImageIdx(updatedList.length - newImages.length);

    if (engineMode !== "RAPID_KEYBOARD") {
      await executeOcrPipeline(newImages, knownSbds, true);
    }
    if (imageInputRef.current) imageInputRef.current.value = "";
  };

  // Rapid Keyboard Entry Handler: Type "8,5" -> automatically saves 8.5 for that student's SBD
  const handleRapidScoreChange = (studentSbd: string, rawVal: string) => {
    const normSbd = normalizeSbdClient(studentSbd);
    if (!normSbd) return;

    setRapidInputs((prev) => ({ ...prev, [normSbd]: rawVal }));

    if (!rawVal.trim()) {
      setOcrRecords((prev) => prev.filter((r) => normalizeSbdClient(r.sbd) !== normSbd));
      setExportResult(null);
      return;
    }

    const parsedScore = parseVietnameseScoreClient(rawVal);
    if (parsedScore === null) return;

    setOcrRecords((prev) => {
      const exists = prev.some((r) => normalizeSbdClient(r.sbd) === normSbd);
      if (exists) {
        return prev.map((r) =>
          normalizeSbdClient(r.sbd) === normSbd
            ? {
                ...r,
                rawScoreText: rawVal,
                score: parsedScore,
                notes: `Nhập nhanh: "${rawVal}" → ${parsedScore}`,
              }
            : r
        );
      }
      return [
        ...prev,
        {
          id: `rapid-${normSbd}-${Date.now()}`,
          sbd: normSbd,
          rawSbd: studentSbd,
          score: parsedScore,
          rawScoreText: rawVal,
          scoreInWords: "",
          confidence: "CAO",
          notes: `Nhập nhanh: "${rawVal}" → ${parsedScore}`,
          sourceImage: "Nhập nhanh bàn phím (Zero-API)",
        },
      ];
    });
    setExportResult(null);
  };

  const handleRapidKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    rowIdx: number,
    totalRows: number
  ) => {
    if (e.key === "Enter" || e.key === "ArrowDown") {
      e.preventDefault();
      const nextIdx = Math.min(totalRows - 1, rowIdx + 1);
      setActiveRapidRowIdx(nextIdx);
      rapidInputRefs.current[nextIdx]?.focus();
      rapidInputRefs.current[nextIdx]?.select();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      const prevIdx = Math.max(0, rowIdx - 1);
      setActiveRapidRowIdx(prevIdx);
      rapidInputRefs.current[prevIdx]?.focus();
      rapidInputRefs.current[prevIdx]?.select();
    }
  };

  const handleApplyBulkPaste = () => {
    if (!bulkPasteText.trim()) return;
    const parsed = parseBulkPastedScores(bulkPasteText, knownSbds);
    if (parsed.length > 0) {
      mergeIncomingRecords(parsed, true);
      setOcrNotice(`Đã dán và chuẩn hóa thành công ${parsed.length} dòng điểm (Không cần API).`);
      setShowBulkPaste(false);
      setBulkPasteText("");
    }
  };

  // ============================================================================
  // STEP 3 HANDLERS: 100% Client-Side Excel Export (Preserves Formatting)
  // ============================================================================
  const triggerBase64Download = (base64Data: string, fileName: string, mimeType: string) => {
    const clean = base64Data.includes(",") ? base64Data.split(",")[1] : base64Data;
    const byteCharacters = atob(clean);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleExecuteAndExport = async () => {
    if (!excelData) {
      setExcelError("Vui lòng tải lên file Excel nguồn ở Bước 1 trước.");
      setCurrentStep(1);
      return;
    }
    setIsExporting(true);
    setExcelError(null);

    try {
      // Runs 100% in the browser using ExcelJS — works on Vercel, GitHub Pages, and offline!
      const result = await exportUpdatedExcelClientSide({
        fileBase64: excelData.fileBase64,
        fileName: excelData.fileName,
        headerRow,
        sbdCol,
        scoreCol,
        sheetIndex: selectedSheetIndex,
        ocrRecords: ocrRecords.map((r) => ({
          sbd: normalizeSbdClient(r.sbd),
          score: r.score,
        })),
      });

      setExportResult(result);
      triggerBase64Download(
        result.updatedExcelBase64,
        result.outputFileName,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
    } catch (err: any) {
      setExcelError(err?.message || "Lỗi khi xuất file Excel.");
    } finally {
      setIsExporting(false);
    }
  };

  const handleDownloadPythonScript = () => {
    const blob = new Blob([liveOpenpyxlScript], { type: "text/x-python;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "auto_update_scores_openpyxl.py";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleCopyCode = () => {
    navigator.clipboard.writeText(liveOpenpyxlScript);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2000);
  };

  const handleUpdateOcrRow = (id: string, field: "sbd" | "rawScoreText", value: string) => {
    setOcrRecords((prev) =>
      prev.map((rec) => {
        if (rec.id !== id) return rec;
        if (field === "sbd") {
          return { ...rec, rawSbd: value, sbd: normalizeSbdClient(value) };
        }
        const parsedScore = parseVietnameseScoreClient(value);
        return {
          ...rec,
          rawScoreText: value,
          score: parsedScore !== null ? parsedScore : rec.score,
          notes:
            parsedScore !== null
              ? `Chuẩn hóa: "${value}" → ${parsedScore}`
              : "Giá trị điểm cần từ 0 đến 10",
        };
      })
    );
    setExportResult(null);
  };

  const handleAddOcrRow = () => {
    const nextIdx = ocrRecords.length + 1;
    const defaultSbd =
      knownSbds[nextIdx - 1] || `K6.${String(nextIdx).padStart(3, "0")}`;
    setOcrRecords((prev) => [
      ...prev,
      {
        id: `manual-${Date.now()}`,
        sbd: defaultSbd,
        rawSbd: defaultSbd,
        score: 8.0,
        rawScoreText: "8,0",
        scoreInWords: "Tám điểm",
        confidence: "CAO",
        notes: "Nhập thủ công",
        sourceImage: "Bổ sung trực tiếp",
      },
    ]);
    setExportResult(null);
  };

  const handleDeleteOcrRow = (id: string) => {
    setOcrRecords((prev) => prev.filter((r) => r.id !== id));
    setExportResult(null);
  };

  const activeImage = images[activeImageIdx] || null;

  return (
    <div className="min-h-screen flex flex-col bg-[#F8FAFC] text-[#0F172A]">
      {/* Top Bar Contract: 3 Zones separated by gap-8 */}
      <header className="sticky top-0 z-30 flex items-center justify-between gap-8 px-6 py-3.5 bg-white border-b border-slate-200">
        <a
          href="#top"
          className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
        >
          VietGrade OCR
        </a>

        {/* 3-Step Sequential Navigation */}
        <div className="flex items-center gap-1.5 p-1 bg-slate-100 rounded-lg">
          <button
            type="button"
            onClick={() => setCurrentStep(1)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors flex items-center gap-2 whitespace-nowrap shrink-0 cursor-pointer ${
              currentStep === 1
                ? "bg-white text-slate-900 shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <span
              className={`w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-mono ${
                excelData
                  ? "bg-emerald-600 text-white"
                  : "bg-slate-300 text-slate-700"
              }`}
            >
              {excelData ? "✓" : "1"}
            </span>
            <span>01. Nạp File Excel Nguồn</span>
          </button>

          <button
            type="button"
            onClick={() => setCurrentStep(2)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors flex items-center gap-2 whitespace-nowrap shrink-0 cursor-pointer ${
              currentStep === 2
                ? "bg-white text-slate-900 shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <span
              className={`w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-mono ${
                ocrRecords.length > 0
                  ? "bg-emerald-600 text-white"
                  : "bg-slate-300 text-slate-700"
              }`}
            >
              {ocrRecords.length > 0 ? "✓" : "2"}
            </span>
            <span>02. Quét Ảnh &amp; Điền Điểm Excel</span>
          </button>
        </div>

        <div className="flex items-center gap-3 shrink-0">
          <input
            ref={excelInputRef}
            type="file"
            accept=".xlsx, .xls"
            onChange={handleExcelUpload}
            className="hidden"
          />
          <button
            type="button"
            onClick={handleExecuteAndExport}
            disabled={isExporting || !excelData || ocrRecords.length === 0}
            className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 flex items-center gap-2 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            <span>{isExporting ? "Đang xử lý..." : "Xuất File Excel Đã Điền Điểm"}</span>
          </button>
        </div>
      </header>

      <main id="top" className="flex-1 max-w-[1400px] w-full mx-auto px-6 py-8 space-y-8">
        {/* Pipeline Status Strip */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 p-4 bg-white border border-slate-200 rounded-lg">
          <div className="flex items-center gap-3.5">
            <div
              className={`p-2.5 rounded-lg ${
                excelData
                  ? "bg-emerald-50 text-emerald-700"
                  : "bg-slate-100 text-slate-500"
              }`}
            >
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <div className="text-xs text-slate-500">
                Bước 1 · File Excel Nguồn
              </div>
              <div className="text-sm font-bold text-slate-900 font-mono">
                {excelData ? excelData.fileName : "Chưa tải lên file Excel nguồn"}
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-6 text-xs font-mono tabular-nums">
            <div>
              <span className="text-slate-500">Học sinh: </span>
              <span className="font-bold text-slate-900">{stats.totalExcelStudents}</span>
            </div>
            <div>
              <span className="text-slate-500">Ảnh phiếu chấm: </span>
              <span className="font-bold text-slate-900">{images.length}</span>
            </div>
            <div>
              <span className="text-slate-500">Đã khớp điểm: </span>
              <span className="font-bold text-emerald-700">
                {stats.matchedCount}/{stats.totalExcelStudents}
              </span>
            </div>
          </div>
        </div>

        {/* =====================================================================
            BƯỚC 1: TẢI LÊN FILE EXCEL NGUỒN TRƯỚC TIÊN (100% CLIENT-SIDE)
           ===================================================================== */}
        {currentStep === 1 && (
          <section className="space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h1 className="text-xl font-bold text-slate-900">
                  01. Tải Lên File Excel Nguồn Cần Nhập Điểm
                </h1>
              </div>
            </div>

            {excelError && (
              <div className="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-center gap-2.5 text-xs text-red-800">
                <AlertCircle className="w-4 h-4 shrink-0 text-red-600" />
                <span>{excelError}</span>
              </div>
            )}

            {/* Drag and Drop Zone */}
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setIsExcelDragging(true);
              }}
              onDragLeave={() => setIsExcelDragging(false)}
              onDrop={handleExcelDrop}
              onClick={() => excelInputRef.current?.click()}
              className={`border-2 border-dashed rounded-xl p-10 transition-colors text-center cursor-pointer ${
                isExcelDragging
                  ? "border-emerald-600 bg-emerald-50/80"
                  : "border-slate-300 hover:border-emerald-600 bg-white"
              }`}
            >
              <div className="max-w-md mx-auto flex flex-col items-center space-y-3">
                <div className="w-12 h-12 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                {excelData && (
                  <div className="text-base font-bold text-slate-900">
                    Đã nạp file nguồn:{" "}
                    <span className="text-emerald-700 font-mono">{excelData.fileName}</span>
                  </div>
                )}
                <div>
                  <span className="px-4 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 rounded-lg hover:bg-emerald-100 transition-colors inline-flex items-center gap-1.5">
                    <Upload className="w-3.5 h-3.5" />
                    <span>
                      {excelData ? "Đổi sang file Excel khác" : "Duyệt chọn file Excel (.xlsx)"}
                    </span>
                  </span>
                </div>
              </div>
            </div>

            {/* Structure Configuration & Roster Preview */}
            {excelData && (
              <div className="space-y-4 bg-white border border-slate-200 rounded-lg p-5">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
                  <div>
                    <h2 className="text-sm font-bold text-slate-900">
                      Cấu Trúc Bảng Điểm Đã Nhận Diện ({excelData.studentRows.length} học sinh)
                    </h2>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Kiểm tra vị trí hàng tiêu đề, cột Số BD và cột Điểm trước khi sang Bước 2.
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={() => setCurrentStep(2)}
                    className="px-4 py-2 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
                  >
                    <span>Tiếp tục: Bước 2 — Tải Ảnh &amp; Nhập Điểm</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
                  {/* Sheet selector */}
                  <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                    <label
                      htmlFor="step1-sheet"
                      className="text-xs font-semibold text-slate-700 block mb-1"
                    >
                      Trang tính (Worksheet):
                    </label>
                    <select
                      id="step1-sheet"
                      value={selectedSheetIndex}
                      onChange={(e) => handleSheetChange(Number(e.target.value))}
                      className="w-full px-2.5 py-1.5 text-xs font-semibold border border-slate-300 rounded bg-white"
                    >
                      {(excelData.availableSheets || [{ index: 0, name: excelData.sheetName, rowCount: excelData.totalRows }]).map(
                        (s) => (
                          <option key={s.index} value={s.index}>
                            {s.name} ({s.rowCount} hàng)
                          </option>
                        )
                      )}
                    </select>
                  </div>

                  <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                    <label
                      htmlFor="step1-header-row"
                      className="text-xs font-semibold text-slate-700 block mb-1"
                    >
                      Hàng tiêu đề (Header Row):
                    </label>
                    <input
                      id="step1-header-row"
                      type="number"
                      min={1}
                      max={20}
                      value={headerRow}
                      onChange={(e) =>
                        handleHeaderRowChange(Math.max(1, Number(e.target.value) || 5))
                      }
                      className="w-full px-2.5 py-1.5 font-mono text-xs font-semibold border border-slate-300 rounded bg-white tabular-nums"
                    />
                  </div>

                  <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                    <label
                      htmlFor="step1-sbd-col"
                      className="text-xs font-semibold text-slate-700 block mb-1"
                    >
                      Cột Số Báo Danh (SBD):
                    </label>
                    <select
                      id="step1-sbd-col"
                      value={sbdCol}
                      onChange={(e) => setSbdCol(e.target.value)}
                      className="w-full px-2.5 py-1.5 font-mono text-xs font-semibold border border-slate-300 rounded bg-white"
                    >
                      {excelData.headers.map((h) => (
                        <option key={h.colLetter} value={h.colLetter}>
                          Cột {h.colLetter} ({h.label})
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="p-3 bg-blue-50/60 rounded-lg border border-blue-200">
                    <label
                      htmlFor="step1-score-col"
                      className="text-xs font-semibold text-blue-900 block mb-1"
                    >
                      Cột Điểm (Chỉ ghi đè cột này):
                    </label>
                    <select
                      id="step1-score-col"
                      value={scoreCol}
                      onChange={(e) => setScoreCol(e.target.value)}
                      className="w-full px-2.5 py-1.5 font-mono text-xs font-bold text-blue-700 border border-blue-300 rounded bg-white"
                    >
                      {excelData.headers.map((h) => (
                        <option key={h.colLetter} value={h.colLetter}>
                          Cột {h.colLetter} ({h.label})
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Roster Preview */}
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-700 flex justify-between items-center">
                    <span>
                      Xem trước danh sách học sinh (Từ hàng {headerRow + 1} trở đi):
                    </span>
                    <span className="font-mono tabular-nums text-slate-500">
                      Tổng cộng: {excelData.studentRows.length} học sinh
                    </span>
                  </div>
                  <div className="overflow-x-auto max-h-[260px]">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 font-semibold text-slate-600">
                        <tr>
                          <th className="py-2 px-3 w-14 text-center">Hàng</th>
                          {excelData.headers.map((h) => (
                            <th key={h.colLetter} className="py-2 px-3 whitespace-nowrap">
                              {h.label} ({h.colLetter})
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {excelData.studentRows.slice(0, 10).map((row) => (
                          <tr key={row.rowNumber} className="hover:bg-slate-50">
                            <td className="py-2 px-3 text-center font-mono tabular-nums text-slate-400">
                              {row.rowNumber}
                            </td>
                            {excelData.headers.map((h) => (
                              <td
                                key={h.colLetter}
                                className={`py-2 px-3 ${
                                  h.colLetter === sbdCol
                                    ? "font-mono font-bold text-slate-900 bg-amber-50/50 tabular-nums"
                                    : "text-slate-700"
                                }`}
                              >
                                {String(row.cells[h.colLetter] ?? "")}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </section>
        )}

        {/* =====================================================================
            BƯỚC 2: TẢI ẢNH PHIẾU CHẤM & TRÍCH XUẤT ĐIỂM (KHÔNG CẦN API / LOCAL WASM)
           ===================================================================== */}
        {currentStep === 2 && (
          <section className="space-y-5">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
              <div>
                <h1 className="text-xl font-bold text-slate-900">
                  02. Tải Ảnh Phiếu Chấm Bài Thi &amp; Trích Xuất Điểm
                </h1>
              </div>

              <input
                ref={imageInputRef}
                type="file"
                accept="image/*"
                multiple
                onChange={handleImageUpload}
                className="hidden"
              />
            </div>

            {/* Optional Bulk Paste Drawer */}
            {showBulkPaste && (
              <div className="p-4 bg-white border border-slate-300 rounded-lg space-y-3">
                <div className="flex items-center justify-between">
                  <div className="text-xs font-bold text-slate-900">
                    Dán Nhanh Danh Sách Điểm
                  </div>
                  <button
                    type="button"
                    onClick={() => setShowBulkPaste(false)}
                    className="text-xs text-slate-500 hover:text-slate-900 cursor-pointer"
                  >
                    Đóng
                  </button>
                </div>
                
                <textarea
                  rows={4}
                  value={bulkPasteText}
                  onChange={(e) => setBulkPasteText(e.target.value)}
                  placeholder={"K6.001  8,5\nK6.002  6,7\nK6.003  7,4"}
                  className="w-full p-2.5 text-xs font-mono border border-slate-300 rounded-lg focus:outline-none focus:border-blue-600"
                />
                <div className="flex justify-end">
                  <button
                    type="button"
                    onClick={handleApplyBulkPaste}
                    className="px-4 py-1.5 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 rounded-lg cursor-pointer"
                  >
                    Áp Dụng Vào Bảng Điểm
                  </button>
                </div>
              </div>
            )}

            {/* Status Notice */}
            {ocrNotice && (
              <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-lg flex items-center gap-2 text-xs text-emerald-900">
                <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-700" />
                <span>{ocrNotice}</span>
              </div>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
              {/* Left Column (5 cols): Grading Sheet Photo Viewer */}
              <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg flex flex-col overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-2 bg-slate-50">
                  <div className="flex items-center gap-2 text-xs font-medium text-slate-700 truncate">
                    <span>Ảnh Phiếu Chấm ({images.length})</span>
                    {activeImage && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="font-mono text-slate-500 truncate">
                          {activeImage.name}
                        </span>
                      </>
                    )}
                  </div>

                  <div className="flex items-center gap-1 shrink-0">
                    <button
                      type="button"
                      onClick={() => setImageZoom((z) => Math.max(0.75, z - 0.25))}
                      className="p-1.5 text-slate-600 hover:text-slate-900 hover:bg-slate-200/60 rounded transition-colors cursor-pointer"
                      title="Thu nhỏ"
                    >
                      <ZoomOut className="w-3.5 h-3.5" />
                    </button>
                    <span className="text-xs font-mono tabular-nums text-slate-600 px-1">
                      {Math.round(imageZoom * 100)}%
                    </span>
                    <button
                      type="button"
                      onClick={() => setImageZoom((z) => Math.min(2.25, z + 0.25))}
                      className="p-1.5 text-slate-600 hover:text-slate-900 hover:bg-slate-200/60 rounded transition-colors cursor-pointer"
                      title="Phóng to kiểm tra nét chữ"
                    >
                      <ZoomIn className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>

                {images.length > 1 && (
                  <div className="px-3 py-2 border-b border-slate-200 flex items-center gap-1.5 overflow-x-auto bg-white">
                    {images.map((img, idx) => (
                      <button
                        key={img.id}
                        type="button"
                        onClick={() => setActiveImageIdx(idx)}
                        className={`px-2.5 py-1 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer ${
                          idx === activeImageIdx
                            ? "bg-slate-900 text-white"
                            : "bg-slate-100 text-slate-700 hover:text-slate-900"
                        }`}
                      >
                        <span>Phiếu #{idx + 1}</span>
                        {img.status === "PROCESSING" && (
                          <RefreshCw className="w-3 h-3 animate-spin text-blue-400" />
                        )}
                        {img.status === "DONE" && (
                          <span
                            className={`text-[10px] font-mono ${
                              idx === activeImageIdx
                                ? "text-emerald-300"
                                : "text-emerald-700"
                            }`}
                          >
                            ✓{img.extractedCount ?? 0}
                          </span>
                        )}
                      </button>
                    ))}
                  </div>
                )}

                <div className="flex-1 min-h-[440px] max-h-[540px] overflow-auto bg-slate-100 p-4 flex items-start justify-center">
                  {activeImage ? (
                    <img
                      src={activeImage.dataUrl}
                      alt={activeImage.name}
                      referrerPolicy="no-referrer"
                      style={{
                        transform: `scale(${imageZoom})`,
                        transformOrigin: "top center",
                      }}
                      className="max-w-full h-auto border border-slate-300 bg-white transition-transform duration-150"
                    />
                  ) : (
                    <div className="my-auto text-center p-8 space-y-4">
                      <ScanText className="w-8 h-8 text-slate-400 mx-auto" />
                      <button
                        type="button"
                        onClick={() => imageInputRef.current?.click()}
                        className="px-4 py-2 text-xs font-bold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
                      >
                        Chọn ảnh phiếu chấm từ máy
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Right Column (7 cols): Either Rapid Keyboard Grid by Excel SBD OR Extracted OCR Table */}
              <div className="lg:col-span-7 bg-white border border-slate-200 rounded-lg flex flex-col overflow-hidden">
                {engineMode === "RAPID_KEYBOARD" && excelData ? (
                  <>
                    <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-4 bg-blue-50/60">
                      <div className="text-xs font-semibold text-blue-950">
                        Nhập Nhanh Theo Danh Sách SBD File Nguồn (Gõ{" "}
                        <code className="font-mono">8,5</code> + nhấn{" "}
                        <code className="font-mono">Enter</code> để tự động xuống dòng)
                      </div>
                      <span className="text-xs font-mono tabular-nums text-blue-800 font-bold">
                        Đã nhập: {stats.matchedCount}/{excelData.studentRows.length}
                      </span>
                    </div>

                    <div className="flex-1 overflow-auto max-h-[540px]">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 text-[11px] font-semibold text-slate-600">
                          <tr>
                            <th className="py-2.5 px-3 w-12 text-center">Hàng</th>
                            <th className="py-2.5 px-3">Số Báo Danh (Cột {sbdCol})</th>
                            <th className="py-2.5 px-3">Họ và tên / Thông tin</th>
                            <th className="py-2.5 px-3 w-36">Gõ Điểm (VD: 8,5)</th>
                            <th className="py-2.5 px-3 text-right w-28">
                              Điểm Chuẩn (Cột {scoreCol})
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {matchedStudentRows.map((row, rIdx) => {
                            const norm = row.normalizedSbd;
                            const inputVal =
                              rapidInputs[norm] ??
                              (row.matchedRawText !== null ? row.matchedRawText : "");
                            const studentName =
                              row.cells["C"] ||
                              row.cells["B"] ||
                              Object.values(row.cells).find(
                                (v) => typeof v === "string" && v !== row.displaySbd
                              ) ||
                              "";

                            return (
                              <tr
                                key={row.rowNumber}
                                className={`transition-colors ${
                                  activeRapidRowIdx === rIdx
                                    ? "bg-blue-50/70"
                                    : "hover:bg-slate-50"
                                }`}
                              >
                                <td className="py-2 px-3 text-center font-mono tabular-nums text-slate-400">
                                  {row.rowNumber}
                                </td>
                                <td className="py-2 px-3 font-mono font-bold text-slate-900 tabular-nums">
                                  {row.displaySbd}
                                </td>
                                <td className="py-2 px-3 text-slate-700 truncate max-w-[200px]">
                                  {String(studentName)}
                                </td>
                                <td className="py-1.5 px-3">
                                  <input
                                    ref={(el) => {
                                      rapidInputRefs.current[rIdx] = el;
                                    }}
                                    type="text"
                                    placeholder="VD: 8,5"
                                    value={inputVal}
                                    onFocus={() => setActiveRapidRowIdx(rIdx)}
                                    onChange={(e) =>
                                      handleRapidScoreChange(row.displaySbd, e.target.value)
                                    }
                                    onKeyDown={(e) =>
                                      handleRapidKeyDown(
                                        e,
                                        rIdx,
                                        matchedStudentRows.length
                                      )
                                    }
                                    className="w-28 px-2.5 py-1 font-mono font-semibold text-slate-900 bg-white border border-slate-300 rounded focus:outline-none focus:border-blue-600 focus:ring-2 focus:ring-blue-100 tabular-nums"
                                  />
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold tabular-nums text-sm">
                                  {row.matchedScore !== null ? (
                                    <span className="text-emerald-700">{row.matchedScore}</span>
                                  ) : (
                                    <span className="text-slate-300 font-normal text-xs">
                                      Trống
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-4 bg-slate-50">
                      <div className="flex items-center gap-2 text-xs font-medium text-slate-700">
                        <span>Bảng Kết Quả OCR (SBD — Điểm Số)</span>
                        <span aria-hidden="true">·</span>
                        <span className="font-mono tabular-nums text-slate-500">
                          {ocrRecords.length} bản ghi
                        </span>
                      </div>

                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setEngineMode("RAPID_KEYBOARD")}
                          className="px-2.5 py-1 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded transition-colors cursor-pointer"
                        >
                          Chuyển sang Gõ Nhanh theo SBD
                        </button>
                        <button
                          type="button"
                          onClick={handleAddOcrRow}
                          className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors flex items-center gap-1 cursor-pointer"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Thêm dòng</span>
                        </button>
                      </div>
                    </div>

                    <div className="flex-1 overflow-auto max-h-[540px]">
                      {isOcrRunning && (
                        <div className="p-4 bg-blue-50/70 border-b border-blue-200 space-y-2 sticky top-0 z-20">
                          <div className="flex items-center justify-between text-xs font-semibold text-blue-900">
                            <div className="flex items-center gap-2">
                              <RefreshCw className="w-3.5 h-3.5 animate-spin text-blue-700 shrink-0" />
                              <span>{ocrProgressText || "Đang xử lý hàng loạt ảnh phiếu chấm..."}</span>
                            </div>
                            <span className="font-mono tabular-nums">{ocrProgressPct}%</span>
                          </div>
                          <div className="w-full h-1.5 bg-blue-100 rounded-full overflow-hidden">
                            <div
                              className="h-full bg-blue-600 transition-all duration-200"
                              style={{ width: `${ocrProgressPct}%` }}
                            />
                          </div>
                        </div>
                      )}

                      {isOcrRunning && ocrRecords.length === 0 ? (
                        <div className="p-6 space-y-2.5">
                          {[1, 2, 3, 4, 5].map((n) => (
                            <div
                              key={n}
                              className="h-9 bg-slate-100 animate-pulse rounded border border-slate-200"
                            />
                          ))}
                        </div>
                      ) : ocrRecords.length === 0 ? (
                        <div className="p-12 text-center space-y-4">
                          <div className="text-sm font-semibold text-slate-700">
                            Chưa có dữ liệu điểm nào
                          </div>
                          <div className="flex items-center justify-center gap-3">
                            <button
                              type="button"
                              onClick={() => setEngineMode("RAPID_KEYBOARD")}
                              className="px-3.5 py-2 text-xs font-semibold text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-lg cursor-pointer"
                            >
                              Gõ Nhanh Theo SBD (1 giây/học sinh)
                            </button>
                          </div>
                        </div>
                      ) : (
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="border-b border-slate-200 bg-slate-50/80 text-[11px] font-semibold text-slate-600">
                              <th className="py-2.5 px-3 w-12 text-center">#</th>
                              <th className="py-2.5 px-3">SBD Chuẩn Hóa</th>
                              <th className="py-2.5 px-3">Điểm Gốc (Dấu phẩy)</th>
                              <th className="py-2.5 px-3 text-right">Điểm Chuẩn (Float)</th>
                              <th className="py-2.5 px-3">Ghi Chú Xác Thực</th>
                              <th className="py-2.5 px-3 w-12 text-center">Xóa</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200 text-xs">
                            {ocrRecords.map((rec, index) => {
                              const isMatchedInExcel = knownSbds.includes(
                                normalizeSbdClient(rec.sbd)
                              );

                              return (
                                <tr
                                  key={rec.id}
                                  className="hover:bg-slate-50/80 transition-colors"
                                >
                                  <td className="py-2 px-3 text-center font-mono tabular-nums text-slate-400">
                                    {index + 1}
                                  </td>
                                  <td className="py-2 px-3">
                                    <div className="flex items-center gap-2">
                                      <input
                                        type="text"
                                        value={rec.sbd}
                                        onChange={(e) =>
                                          handleUpdateOcrRow(rec.id, "sbd", e.target.value)
                                        }
                                        className="w-28 px-2 py-1 font-mono font-semibold text-slate-900 bg-slate-50 border border-slate-300 rounded focus:outline-none focus:border-blue-600 focus:bg-white tabular-nums"
                                      />
                                      {isMatchedInExcel ? (
                                        <span className="text-[11px] text-emerald-700 font-medium whitespace-nowrap">
                                          ✓ Khớp Excel
                                        </span>
                                      ) : (
                                        <span className="text-[11px] text-amber-700 font-medium whitespace-nowrap">
                                          ! Ngoài DS
                                        </span>
                                      )}
                                    </div>
                                  </td>
                                  <td className="py-2 px-3">
                                    <input
                                      type="text"
                                      value={rec.rawScoreText}
                                      onChange={(e) =>
                                        handleUpdateOcrRow(
                                          rec.id,
                                          "rawScoreText",
                                          e.target.value
                                        )
                                      }
                                      className="w-20 px-2 py-1 font-mono text-slate-700 bg-slate-50 border border-slate-300 rounded focus:outline-none focus:border-blue-600 focus:bg-white tabular-nums"
                                    />
                                  </td>
                                  <td className="py-2 px-3 text-right font-mono font-bold text-blue-700 tabular-nums text-sm">
                                    {rec.score}
                                  </td>
                                  <td className="py-2 px-3 text-slate-600 truncate max-w-[220px]">
                                    {rec.notes}
                                  </td>
                                  <td className="py-2 px-3 text-center">
                                    <button
                                      type="button"
                                      onClick={() => handleDeleteOcrRow(rec.id)}
                                      className="p-1 text-slate-400 hover:text-red-600 rounded transition-colors cursor-pointer"
                                      title="Xóa dòng này"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      )}
                    </div>
                  </>
                )}
              </div>
            </div>

            {/* Live Matched Excel Preview (Hình 3: Kết Quả So Khớp SBD Hình 1 × Hình 2) */}
            {excelData && ocrRecords.length > 0 && (
              <div className="bg-white border-2 border-emerald-600/80 rounded-lg overflow-hidden shadow-xs">
                <div className="px-4 py-3 bg-emerald-50/80 border-b border-emerald-200 flex items-center justify-between">
                  <div className="text-xs font-bold text-emerald-950 flex items-center gap-2">
                    <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                    <span>
                      KẾT QUẢ SO KHỚP CỘT SBD: Đã khớp{" "}
                      {stats.matchedCount}/{stats.totalExcelStudents} học sinh
                    </span>
                  </div>
                </div>

                {/* Excel Sheet Title Header Reproduction (Rows 1-4) */}
                <div className="px-5 py-3 bg-white border-b border-slate-200 text-xs font-serif text-slate-900">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-1">
                    <div>
                      <div className="text-[13px] tracking-tight">
                        SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH NINH BÌNH
                      </div>
                      <div className="text-[13px] font-bold pl-8">TRƯỜNG THCS MỸ XÁ</div>
                    </div>
                  </div>
                  <div className="text-center font-bold text-sm mt-1">
                    BẢNG ĐIỂM THI MÔN: VĂN 6
                  </div>
                </div>

                {/* Matched Excel Table (Exact Hình 3 Layout) */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="text-[11px] font-mono text-slate-500 bg-slate-100 border-b border-slate-300">
                        <th className="py-1 px-2 text-center border-r border-slate-300 w-10">
                          Hàng
                        </th>
                        {excelData.headers.map((h) => (
                          <th
                            key={h.colLetter}
                            className={`py-1 px-2.5 text-center border-r border-slate-300 ${
                              h.colLetter === sbdCol
                                ? "bg-amber-100 text-amber-900 font-bold"
                                : h.colLetter === scoreCol
                                ? "bg-emerald-100 text-emerald-900 font-bold"
                                : ""
                            }`}
                          >
                            {h.colLetter}
                          </th>
                        ))}
                      </tr>
                      <tr className="bg-[#D9D9D9] text-slate-900 font-bold border-b border-slate-400">
                        <th className="py-2 px-2 text-center font-mono text-[11px] text-slate-600 bg-slate-100 border-r border-slate-300">
                          {headerRow}
                        </th>
                        {excelData.headers.map((h) => (
                          <th
                            key={h.colLetter}
                            className="py-2 px-3 text-center border-r border-slate-400 whitespace-nowrap"
                          >
                            {h.label}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-300">
                      {matchedStudentRows.map((row) => (
                        <tr key={row.rowNumber} className="hover:bg-slate-50">
                          <td className="py-2 px-2 text-center font-mono text-[11px] text-slate-500 bg-slate-50 border-r border-slate-300 tabular-nums">
                            {row.rowNumber}
                          </td>
                          {excelData.headers.map((h) => {
                            const isScore = h.colLetter === scoreCol;
                            const isSbd = h.colLetter === sbdCol;
                            const isName = h.colLetter === "B";
                            const val = row.cells[h.colLetter];

                            if (isScore) {
                              return (
                                <td
                                  key={h.colLetter}
                                  className={`py-2 px-3 text-center font-mono tabular-nums border-r border-slate-300 ${
                                    row.isMatched
                                      ? "bg-emerald-50/70 font-bold text-slate-900"
                                      : "text-slate-400"
                                  }`}
                                >
                                  {row.isMatched ? row.matchedScore : ""}
                                </td>
                              );
                            }

                            return (
                              <td
                                key={h.colLetter}
                                className={`py-2 px-3 border-r border-slate-300 ${
                                  isSbd
                                    ? "text-center font-mono font-semibold text-slate-900 bg-amber-50/30 tabular-nums"
                                    : isName
                                    ? "text-left text-slate-900"
                                    : "text-center text-slate-800 tabular-nums"
                                }`}
                              >
                                {val !== null && val !== undefined ? String(val) : ""}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Excel Export Card Directly In Step 2 */}
            {excelData && ocrRecords.length > 0 && (
              <div className="max-w-2xl mx-auto w-full bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-5">
                <div className="flex items-center justify-between">
                  <div>
                    <h2 className="text-base font-bold text-slate-900">
                      Xuất File Excel Đã Nhập Điểm Hoàn Chỉnh
                    </h2>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Đã điền điểm số thực Float vào cột {scoreCol}, bảo toàn 100% định dạng file nguồn.
                    </p>
                  </div>
                  <div className="w-10 h-10 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700">
                    <FileCheck2 className="w-5 h-5" />
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-3 text-xs border-t border-b border-slate-200 py-3.5 bg-slate-50/60 rounded-lg px-3">
                  <div className="text-center">
                    <div className="text-slate-500 text-[11px]">Tiêu đề Sở/Trường</div>
                    <div className="font-mono text-emerald-700 font-bold mt-0.5">
                      Bảo toàn ({excelData?.mergedCellsCount || 0} Merges)
                    </div>
                  </div>
                  <div className="text-center border-x border-slate-200">
                    <div className="text-slate-500 text-[11px]">Font chữ &amp; Viền kẻ</div>
                    <div className="font-mono text-slate-900 font-bold mt-0.5">Giữ nguyên 100%</div>
                  </div>
                  <div className="text-center">
                    <div className="text-slate-500 text-[11px]">Kiểu dữ liệu điểm</div>
                    <div className="font-mono text-blue-700 font-bold mt-0.5">Số thực (Float)</div>
                  </div>
                </div>

                {exportResult && (
                  <pre className="p-3 bg-slate-900 text-emerald-300 font-mono text-[11px] rounded-lg border border-slate-800 overflow-x-auto whitespace-pre-wrap leading-relaxed">
                    {exportResult.pythonAuditLog}
                  </pre>
                )}

                <button
                  type="button"
                  onClick={handleExecuteAndExport}
                  disabled={isExporting || !excelData || ocrRecords.length === 0}
                  className="w-full py-3 px-4 text-sm font-bold text-white bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 rounded-lg transition-colors flex items-center justify-center gap-2 shadow-xs cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>
                    {isExporting
                      ? "Đang xuất file Excel..."
                      : exportResult
                      ? `Tải Lại ${exportResult.outputFileName}`
                      : "Tải File Excel Đã Nhập Điểm (.xlsx)"}
                  </span>
                </button>
              </div>
            )}
          </section>
        )}
      </main>

      <footer className="mt-12 border-t border-slate-200 bg-white py-4 px-6">
        <div className="max-w-[1400px] mx-auto flex items-center justify-center text-xs font-semibold text-slate-700 tracking-wide">
          <span>VIETLEQUOC117@GMAIL.COM : 0837110784</span>
        </div>
      </footer>
    </div>
  );
}
