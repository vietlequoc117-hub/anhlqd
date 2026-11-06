import React, { useState, useMemo, useRef } from "react";
import {
  Upload,
  FileSpreadsheet,
  ScanText,
  Download,
  CheckCircle2,
  AlertCircle,
  Plus,
  Trash2,
  RefreshCw,
  ZoomIn,
  ZoomOut,
  FileCheck2,
  ArrowRight,
  Link2,
  Sparkles,
} from "lucide-react";
import {
  OcrRecord,
  ExcelSheetData,
  normalizeSbdClient,
  canonicalSbdKey,
  parseVietnameseScoreClient,
  generateSampleGradingSheetImage,
} from "./utils/sampleGradingSheet";
import {
  parseExcelClientSide,
  detectExcelFileRole,
  exportUpdatedExcelClientSide,
  generateSampleAnh9PhachBundleClientSide,
} from "./utils/clientExcelEngine";
import {
  runLocalBrowserOcr,
  parseBulkPastedScores,
} from "./utils/localOcrEngine";
import { convertPdfFileToImages } from "./utils/pdfToImages";

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
  // Sequential Step Navigation (1: Nạp 2 File Excel Nguồn -> 2: Quét PDF/Ảnh & Xuất Excel Hình 4)
  const [currentStep, setCurrentStep] = useState<1 | 2>(1);

  // STEP 1A: File Excel Nguồn 1 (Hình 1: Bảng Điểm Thí Sinh - có cột SBD & cột Điểm)
  const [excelData, setExcelData] = useState<ExcelSheetData | null>(null);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [headerRow, setHeaderRow] = useState<number>(6);
  const [sbdCol, setSbdCol] = useState<string>("G");
  const [scoreCol, setScoreCol] = useState<string>("I");

  // STEP 1B: File Excel Nguồn 2 (Hình 2: Bảng Ghép Mã Phách - có cột SBD & cột MÃ PHÁCH)
  const [phachExcelData, setPhachExcelData] = useState<ExcelSheetData | null>(null);
  const [phachSheetIndex, setPhachSheetIndex] = useState<number>(0);
  const [phachHeaderRow, setPhachHeaderRow] = useState<number>(4);
  const [phachSbdCol, setPhachSbdCol] = useState<string>("B");
  const [phachCodeCol, setPhachCodeCol] = useState<string>("C");

  const [isExcelLoading, setIsExcelLoading] = useState<boolean>(false);
  const [excelError, setExcelError] = useState<string | null>(null);
  const [isExcelDragging1, setIsExcelDragging1] = useState<boolean>(false);
  const [isExcelDragging2, setIsExcelDragging2] = useState<boolean>(false);

  // STEP 2: Grading Sheet PDF/Images (Hình 3) & Hybrid OCR State
  const [engineMode, setEngineMode] = useState<OcrEngineMode>("HYBRID_CLOUD");
  const [images, setImages] = useState<UploadedImage[]>([]);
  const [activeImageIdx, setActiveImageIdx] = useState<number>(0);
  const [imageZoom, setImageZoom] = useState<number>(1);
  const [isOcrRunning, setIsOcrRunning] = useState<boolean>(false);
  const [ocrProgressText, setOcrProgressText] = useState<string>("");
  const [ocrProgressPct, setOcrProgressPct] = useState<number>(0);
  const [ocrNotice, setOcrNotice] = useState<string | null>(null);
  const [ocrRecords, setOcrRecords] = useState<OcrRecord[]>([]);

  // Rapid Keyboard Entry & Bulk Paste State
  const [rapidInputs, setRapidInputs] = useState<Record<string, string>>({});
  const [activeRapidRowIdx, setActiveRapidRowIdx] = useState<number>(0);
  const [showBulkPaste, setShowBulkPaste] = useState<boolean>(false);
  const [bulkPasteText, setBulkPasteText] = useState<string>("");
  const rapidInputRefs = useRef<Record<number, HTMLInputElement | null>>({});

  // STEP 3: Export State (Hình 4)
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [exportResult, setExportResult] = useState<{
    updatedExcelBase64: string;
    outputFileName: string;
    matchedCount: number;
    untouchedCount: number;
    pythonAuditLog: string;
    openpyxlCode: string;
  } | null>(null);

  const excel1InputRef = useRef<HTMLInputElement>(null);
  const excel2InputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);

  // Build bidirectional SBD <-> MÃ PHÁCH dictionaries from File Excel 2 (Hình 2)
  const { sbdToPhachMap, phachToSbdMap, totalPhachMappedCount } = useMemo(() => {
    const s2p: Record<string, string> = {};
    const p2s: Record<string, string> = {};
    let count = 0;

    if (!phachExcelData) {
      return { sbdToPhachMap: s2p, phachToSbdMap: p2s, totalPhachMappedCount: 0 };
    }

    // First index all cross-sheet pairs if present
    if (
      Array.isArray(phachExcelData.allSheetsPhachPairs) &&
      phachExcelData.allSheetsPhachPairs.length > 0
    ) {
      for (const pair of phachExcelData.allSheetsPhachPairs) {
        const normSbd = normalizeSbdClient(pair.sbd);
        const canonSbd = canonicalSbdKey(pair.sbd);
        const normPhach = normalizeSbdClient(pair.maPhach);
        const canonPhach = canonicalSbdKey(pair.maPhach);

        if (normSbd && normPhach) {
          s2p[normSbd] = normPhach;
          if (canonSbd) s2p[canonSbd] = normPhach;
          p2s[normPhach] = normSbd;
          if (canonPhach) p2s[canonPhach] = normSbd;
        }
      }
    }

    // Override/ensure current sheet's selected columns (phachSbdCol & phachCodeCol) are indexed accurately
    for (const row of phachExcelData.studentRows) {
      const rawSbd = row.cells[phachSbdCol] ?? row.sbd;
      const rawPhach = row.cells[phachCodeCol] ?? row.maPhach;
      const normSbd = normalizeSbdClient(rawSbd);
      const canonSbd = canonicalSbdKey(rawSbd);
      const normPhach = normalizeSbdClient(rawPhach);
      const canonPhach = canonicalSbdKey(rawPhach);

      if (normSbd) {
        s2p[normSbd] = normPhach || "";
        if (canonSbd) s2p[canonSbd] = normPhach || "";
      }
      if (normSbd && normPhach) {
        p2s[normPhach] = normSbd;
        if (canonPhach) p2s[canonPhach] = normSbd;
        count++;
      }
    }

    return {
      sbdToPhachMap: s2p,
      phachToSbdMap: p2s,
      totalPhachMappedCount: count,
    };
  }, [phachExcelData, phachSbdCol, phachCodeCol]);

  // Extract list of known SBDs from File Excel 1 (Hình 1)
  const knownSbds = useMemo(() => {
    if (!excelData) return [];
    return excelData.studentRows
      .map((r) => normalizeSbdClient(r.cells[sbdCol] ?? r.sbd))
      .filter(Boolean);
  }, [excelData, sbdCol]);

  // Dynamically resolve each OCR record through phachToSbdMap & sbdToPhachMap
  // so whether the grading sheet (Hình 3) has MÃ PHÁCH (A40, A41...) or SBD (K9.001...),
  // both `resolvedSbd` and `resolvedPhach` are always populated!
  const resolvedOcrRecords = useMemo(() => {
    return ocrRecords.map((rec) => {
      const normCode = normalizeSbdClient(rec.sbd);
      const canonCode = canonicalSbdKey(rec.sbd);
      const normExplicitPhach = normalizeSbdClient(rec.maPhach || "");
      const canonExplicitPhach = canonicalSbdKey(rec.maPhach || "");

      let finalSbd = normCode;
      let finalPhach = normExplicitPhach;

      // Case 1: The code in rec.sbd is actually a MÃ PHÁCH in phachToSbdMap (e.g., "A40" -> "K9.001")
      if (phachToSbdMap[normCode] || phachToSbdMap[canonCode]) {
        finalPhach = normCode;
        finalSbd = phachToSbdMap[normCode] || phachToSbdMap[canonCode];
      } else if (
        normExplicitPhach &&
        (phachToSbdMap[normExplicitPhach] || phachToSbdMap[canonExplicitPhach])
      ) {
        // Case 2: rec.maPhach is in phachToSbdMap
        finalSbd =
          phachToSbdMap[normExplicitPhach] || phachToSbdMap[canonExplicitPhach];
      } else if (sbdToPhachMap[normCode] || sbdToPhachMap[canonCode]) {
        // Case 3: rec.sbd is already an SBD (e.g. "K9.001"), look up its Mã Phách ("A40")
        finalPhach = sbdToPhachMap[normCode] || sbdToPhachMap[canonCode];
      }

      return {
        ...rec,
        resolvedSbd: finalSbd,
        resolvedPhach: finalPhach,
      };
    });
  }, [ocrRecords, phachToSbdMap, sbdToPhachMap]);

  // Lookup map for fast 3-way VLOOKUP matching (indexes by SBD, canonical SBD, Mã Phách, and canonical Mã Phách)
  const ocrLookupMap = useMemo(() => {
    const map = new Map<string, (typeof resolvedOcrRecords)[number]>();
    for (const rec of resolvedOcrRecords) {
      const normS = normalizeSbdClient(rec.resolvedSbd);
      const canonS = canonicalSbdKey(rec.resolvedSbd);
      if (normS) map.set(normS, rec);
      if (canonS) map.set(canonS, rec);

      if (rec.resolvedPhach) {
        const normP = normalizeSbdClient(rec.resolvedPhach);
        const canonP = canonicalSbdKey(rec.resolvedPhach);
        if (normP) map.set(normP, rec);
        if (canonP) map.set(canonP, rec);
      }
    }
    return map;
  }, [resolvedOcrRecords]);

  // Match File Excel 1 (Hình 1) student rows with File Excel 2 (Hình 2 Mã Phách) & Hình 3 OCR Scores -> Hình 4!
  const matchedStudentRows = useMemo(() => {
    if (!excelData) return [];
    return excelData.studentRows.map((row) => {
      const rawSbd = row.cells[sbdCol] ?? row.sbd;
      const normSbd = normalizeSbdClient(rawSbd);
      const canonSbd = canonicalSbdKey(rawSbd);

      const linkedPhach =
        (normSbd ? sbdToPhachMap[normSbd] : undefined) ??
        (canonSbd ? sbdToPhachMap[canonSbd] : undefined) ??
        "";
      const normPhach = normalizeSbdClient(linkedPhach);
      const canonPhach = canonicalSbdKey(linkedPhach);

      const matchedOcr =
        (normSbd ? ocrLookupMap.get(normSbd) : undefined) ||
        (canonSbd ? ocrLookupMap.get(canonSbd) : undefined) ||
        (normPhach ? ocrLookupMap.get(normPhach) : undefined) ||
        (canonPhach ? ocrLookupMap.get(canonPhach) : undefined);

      return {
        ...row,
        displaySbd: rawSbd !== null && rawSbd !== undefined ? String(rawSbd) : "",
        normalizedSbd: normSbd,
        linkedPhach,
        matchedScore: matchedOcr ? matchedOcr.score : null,
        matchedRawText: matchedOcr ? matchedOcr.rawScoreText : null,
        matchedSource: matchedOcr ? matchedOcr.sourceImage : null,
        isMatched: Boolean(matchedOcr),
      };
    });
  }, [excelData, sbdCol, sbdToPhachMap, ocrLookupMap]);

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

  // ============================================================================
  // STEP 1 HANDLERS: Load File Excel 1 (Hình 1) & File Excel 2 (Hình 2)
  // ============================================================================
  const readBlobAsBase64 = (file: File | Blob): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });

  const loadRosterExcelFile = async (
    file: File | Blob,
    fileName: string,
    sheetIndex = 0
  ) => {
    const base64 = await readBlobAsBase64(file);
    const parsed = await parseExcelClientSide({
      fileBase64: base64,
      fileName,
      targetSheetIndex: sheetIndex,
    });
    setExcelData(parsed);
    setSelectedSheetIndex(sheetIndex);
    setHeaderRow(parsed.detectedHeaderRow || 6);
    setSbdCol(parsed.detectedSbdCol || "G");
    setScoreCol(parsed.detectedScoreCol || "I");
    setExportResult(null);
  };

  const loadPhachExcelFile = async (
    file: File | Blob,
    fileName: string,
    sheetIndex = 0
  ) => {
    const base64 = await readBlobAsBase64(file);
    const parsed = await parseExcelClientSide({
      fileBase64: base64,
      fileName,
      targetSheetIndex: sheetIndex,
    });
    setPhachExcelData(parsed);
    setPhachSheetIndex(sheetIndex);
    setPhachHeaderRow(parsed.detectedHeaderRow || 4);
    setPhachSbdCol(parsed.detectedSbdCol || "B");
    setPhachCodeCol(parsed.detectedPhachCol || "C");
    setExportResult(null);
  };

  /**
   * Handles uploading 1 or 2 Excel files at once.
   * If 2 files are selected simultaneously, automatically assigns the one with "MÃ PHÁCH" to Hình 2
   * and the student roster to Hình 1!
   */
  const handleSmartExcelFilesUpload = async (
    fileList: FileList | File[],
    targetSlot: "SLOT_1_ROSTER" | "SLOT_2_PHACH"
  ) => {
    const files = Array.from(fileList).filter(
      (f) => f.name.endsWith(".xlsx") || f.name.endsWith(".xls")
    );
    if (files.length === 0) {
      setExcelError("Vui lòng chọn file Excel có định dạng .xlsx.");
      return;
    }

    setIsExcelLoading(true);
    setExcelError(null);

    try {
      if (files.length >= 2) {
        // User selected both Excel files at once! Automatically classify Hình 1 vs Hình 2
        for (const f of files) {
          const base64 = await readBlobAsBase64(f);
          const parsed = await parseExcelClientSide({
            fileBase64: base64,
            fileName: f.name,
            targetSheetIndex: 0,
          });
          const role = detectExcelFileRole(parsed);
          if (role === "PHACH_MAP") {
            setPhachExcelData(parsed);
            setPhachSheetIndex(0);
            setPhachHeaderRow(parsed.detectedHeaderRow || 4);
            setPhachSbdCol(parsed.detectedSbdCol || "B");
            setPhachCodeCol(parsed.detectedPhachCol || "C");
          } else {
            setExcelData(parsed);
            setSelectedSheetIndex(0);
            setHeaderRow(parsed.detectedHeaderRow || 6);
            setSbdCol(parsed.detectedSbdCol || "G");
            setScoreCol(parsed.detectedScoreCol || "I");
          }
        }
        setExportResult(null);
      } else {
        const singleFile = files[0];
        if (targetSlot === "SLOT_2_PHACH") {
          await loadPhachExcelFile(singleFile, singleFile.name, 0);
        } else {
          await loadRosterExcelFile(singleFile, singleFile.name, 0);
        }
      }
    } catch (err: any) {
      setExcelError(
        err?.message ||
          "Không thể đọc file Excel. Vui lòng đảm bảo file có định dạng .xlsx chuẩn."
      );
    } finally {
      setIsExcelLoading(false);
      if (excel1InputRef.current) excel1InputRef.current.value = "";
      if (excel2InputRef.current) excel2InputRef.current.value = "";
    }
  };

  const handleSheetChange1 = async (newSheetIdx: number) => {
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
      setHeaderRow(parsed.detectedHeaderRow || 6);
      setSbdCol(parsed.detectedSbdCol || "G");
      setScoreCol(parsed.detectedScoreCol || "I");
    } catch (err: any) {
      setExcelError(err?.message || "Không thể chuyển trang tính.");
    } finally {
      setIsExcelLoading(false);
    }
  };

  const handleHeaderRowChange1 = async (newHeaderRow: number) => {
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

  const handleSheetChange2 = async (newSheetIdx: number) => {
    if (!phachExcelData) return;
    setPhachSheetIndex(newSheetIdx);
    setIsExcelLoading(true);
    try {
      const parsed = await parseExcelClientSide({
        fileBase64: phachExcelData.fileBase64,
        fileName: phachExcelData.fileName,
        targetSheetIndex: newSheetIdx,
      });
      setPhachExcelData(parsed);
      setPhachHeaderRow(parsed.detectedHeaderRow || 4);
      setPhachSbdCol(parsed.detectedSbdCol || "B");
      setPhachCodeCol(parsed.detectedPhachCol || "C");
    } catch (err: any) {
      setExcelError(err?.message || "Không thể chuyển trang tính File 2.");
    } finally {
      setIsExcelLoading(false);
    }
  };

  const handleHeaderRowChange2 = async (newHeaderRow: number) => {
    setPhachHeaderRow(newHeaderRow);
    if (!phachExcelData) return;
    try {
      const parsed = await parseExcelClientSide({
        fileBase64: phachExcelData.fileBase64,
        fileName: phachExcelData.fileName,
        preferredHeaderRow: newHeaderRow,
        preferredSbdCol: phachSbdCol,
        preferredPhachCol: phachCodeCol,
        targetSheetIndex: phachSheetIndex,
      });
      setPhachExcelData(parsed);
    } catch {
      // Keep current state
    }
  };

  // Load Sample Workspace: Hình 1 (Bảng điểm Anh 9) + Hình 2 (Bảng Mã Phách) + Hình 3 (Phiếu chấm Mã Phách) -> Hình 4
  const handleLoadSamplePhachWorkspace = async () => {
    setIsExcelLoading(true);
    setOcrNotice(null);
    setExcelError(null);

    try {
      const { rosterExcel, phachExcel } =
        await generateSampleAnh9PhachBundleClientSide();

      setExcelData(rosterExcel);
      setSelectedSheetIndex(0);
      setHeaderRow(rosterExcel.detectedHeaderRow || 6);
      setSbdCol(rosterExcel.detectedSbdCol || "G");
      setScoreCol(rosterExcel.detectedScoreCol || "I");

      setPhachExcelData(phachExcel);
      setPhachSheetIndex(0);
      setPhachHeaderRow(phachExcel.detectedHeaderRow || 4);
      setPhachSbdCol(phachExcel.detectedSbdCol || "B");
      setPhachCodeCol(phachExcel.detectedPhachCol || "C");

      // Build phachToSbdMap from sample phachExcel
      const sampleP2S: Record<string, string> = {};
      for (const r of phachExcel.studentRows) {
        const s = normalizeSbdClient(r.sbd);
        const p = normalizeSbdClient(r.maPhach);
        if (s && p) {
          sampleP2S[p] = s;
        }
      }

      // Load Hình 3 (/assets/aistudio/image_a4899c97.png)
      let dataUrl = "";
      try {
        const resp = await fetch("/assets/aistudio/image_a4899c97.png");
        if (resp.ok) {
          const blob = await resp.blob();
          dataUrl = await readBlobAsBase64(blob);
        }
      } catch {
        // fallback
      }
      if (!dataUrl) {
        dataUrl = generateSampleGradingSheetImage().dataUrl;
      }

      const sampleImageObj: UploadedImage = {
        id: `hinh3-${Date.now()}`,
        name: "Hinh_3_Phieu_Cham_Ma_Phach_Anh_9.png",
        mimeType: "image/png",
        dataUrl,
      };
      setImages([sampleImageObj]);
      setActiveImageIdx(0);

      const sbdsFromSample = rosterExcel.studentRows
        .map((r) => normalizeSbdClient(r.cells["G"] ?? r.sbd))
        .filter(Boolean);

      setCurrentStep(2);
      await executeOcrPipeline(
        [sampleImageObj],
        sbdsFromSample,
        sampleP2S,
        false
      );
    } catch (err: any) {
      setExcelError(err?.message || "Lỗi khi khởi tạo bộ dữ liệu mẫu.");
    } finally {
      setIsExcelLoading(false);
    }
  };

  // ============================================================================
  // STEP 2 HANDLERS: Hybrid Cloud OCR / Local OCR / Rapid Keyboard
  // ============================================================================
  const mergeIncomingRecords = (
    incoming: OcrRecord[],
    activePhachMap: Record<string, string>,
    append: boolean
  ) => {
    setOcrRecords((prev) => {
      const combined = append ? [...prev, ...incoming] : incoming;
      const byKey = new Map<string, OcrRecord>();

      for (const item of combined) {
        const rawCode = normalizeSbdClient(item.sbd);
        const canonCode = canonicalSbdKey(item.sbd);
        const rawPhach = normalizeSbdClient(item.maPhach || "");
        const canonPhach = canonicalSbdKey(item.maPhach || "");

        let mappedSbd = rawCode;
        let mappedPhach = rawPhach;

        if (activePhachMap[rawCode] || activePhachMap[canonCode]) {
          mappedPhach = rawCode;
          mappedSbd = activePhachMap[rawCode] || activePhachMap[canonCode];
        } else if (
          rawPhach &&
          (activePhachMap[rawPhach] || activePhachMap[canonPhach])
        ) {
          mappedSbd = activePhachMap[rawPhach] || activePhachMap[canonPhach];
        }

        const key =
          canonicalSbdKey(mappedSbd) ||
          normalizeSbdClient(mappedSbd) ||
          canonicalSbdKey(mappedPhach);

        if (key) {
          byKey.set(key, {
            ...item,
            sbd: mappedSbd,
            maPhach: mappedPhach || item.maPhach,
          });
        }
      }

      const nextList = Array.from(byKey.values());
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
    activePhachToSbdMap: Record<string, string>,
    append = false
  ) => {
    if (targetImages.length === 0) return;

    if (engineMode === "RAPID_KEYBOARD") {
      setOcrNotice(
        "Đang ở Chế độ Nhập Nhanh Bàn Phím: Gõ điểm (VD: 8,5) và nhấn Enter để xuống dòng tức thì."
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
    let missingCloudKeyMessage: string | null = null;

    // Combine known SBDs and known Mã Phách codes so OCR knows all valid identifiers
    const combinedKnownCodes = Array.from(
      new Set([...activeKnownSbds, ...Object.keys(activePhachToSbdMap)])
    );

    try {
      for (let idx = 0; idx < totalFiles; idx++) {
        const currentImg = targetImages[idx];
        const basePct = Math.round((idx / totalFiles) * 100);

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
            phachToSbdMap: activePhachToSbdMap,
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
          let cloudSuccess = false;
          for (let attempt = 1; attempt <= 3; attempt++) {
            try {
              const response = await fetch("/api/ocr", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  knownSbds: combinedKnownCodes,
                  phachToSbdMap: activePhachToSbdMap,
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
                let errPayload: any = null;
                try {
                  errPayload = await response.json();
                } catch {
                  // ignore non-json error body
                }
                if (
                  response.status === 503 &&
                  errPayload?.errorCode === "MISSING_GEMINI_API_KEY"
                ) {
                  missingCloudKeyMessage =
                    "Lưu ý trên Vercel: Để AI đọc chữ viết tay bút đỏ chính xác 100% như trên Google AI Studio, bạn cần vào Vercel -> Project Settings -> Environment Variables -> Thêm biến GEMINI_API_KEY rồi bấm Redeploy.";
                  break;
                }
                throw new Error(errPayload?.error || `HTTP ${response.status}`);
              }
              const data = await response.json();
              if (Array.isArray(data.records) && data.records.length > 0) {
                sheetRecords = data.records;
                cloudSuccess = true;
                break;
              } else {
                throw new Error("Empty cloud records");
              }
            } catch {
              if (attempt < 3) {
                setOcrProgressText(
                  `[Phiếu ${idx + 1}/${totalFiles}] Đang thử lại kết nối AI OCR (lần ${attempt + 1}/3)...`
                );
                await new Promise((r) => setTimeout(r, 700 * attempt));
              }
            }
          }

          if (!cloudSuccess) {
            sheetRecords = await runLocalBrowserOcr({
              images: [currentImg],
              knownSbds: activeKnownSbds,
              phachToSbdMap: activePhachToSbdMap,
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

        if (sheetRecords.length > 0) {
          mergeIncomingRecords(sheetRecords, activePhachToSbdMap, true);
        }

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

      if (missingCloudKeyMessage) {
        setOcrNotice(missingCloudKeyMessage);
      } else {
        const phachNote =
          Object.keys(activePhachToSbdMap).length > 0
            ? ` và tự động quy chiếu cột MÃ PHÁCH → SBD`
            : "";
        setOcrNotice(
          `Đã quét hoàn tất ${totalFiles} trang phiếu chấm, trích xuất ${totalExtractedAcrossBatch} dòng điểm${phachNote}.`
        );
      }
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

    setOcrNotice(`Đang xử lý & nạp ${files.length} tệp phiếu chấm (Ảnh / PDF)...`);
    const newImages: UploadedImage[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const isPdf =
        file.type === "application/pdf" || /\.pdf$/i.test(file.name);

      if (isPdf) {
        try {
          const pdfPages = await convertPdfFileToImages(
            file,
            (pageNum, totalPages) => {
              setOcrNotice(
                `Đang tách trang ${pageNum}/${totalPages} từ file PDF "${file.name}"...`
              );
            }
          );
          pdfPages.forEach((p, pageIdx) => {
            newImages.push({
              id: `${file.name}-p${pageIdx + 1}-${Date.now()}-${i}`,
              name: p.name,
              mimeType: p.mimeType,
              dataUrl: p.dataUrl,
              status: "PENDING",
            });
          });
        } catch (err: any) {
          setOcrNotice(
            `Không thể đọc file PDF "${file.name}": ${
              err?.message || "Định dạng PDF không hợp lệ."
            }`
          );
        }
      } else {
        const optimized = await optimizeImageForBatchOcr(file);
        newImages.push({
          id: `${file.name}-${Date.now()}-${i}`,
          name: file.name,
          mimeType: optimized.mimeType,
          dataUrl: optimized.dataUrl,
          status: "PENDING",
        });
      }
    }

    if (newImages.length === 0) {
      if (imageInputRef.current) imageInputRef.current.value = "";
      return;
    }

    const updatedList = [...images, ...newImages];
    setImages(updatedList);
    setActiveImageIdx(updatedList.length - newImages.length);

    if (engineMode !== "RAPID_KEYBOARD") {
      await executeOcrPipeline(newImages, knownSbds, phachToSbdMap, true);
    }
    if (imageInputRef.current) imageInputRef.current.value = "";
  };

  // Rapid Keyboard Entry Handler
  const handleRapidScoreChange = (
    studentSbd: string,
    linkedPhach: string,
    rawVal: string
  ) => {
    const normSbd = normalizeSbdClient(studentSbd);
    if (!normSbd) return;

    setRapidInputs((prev) => ({ ...prev, [normSbd]: rawVal }));

    if (!rawVal.trim()) {
      setOcrRecords((prev) =>
        prev.filter((r) => normalizeSbdClient(r.sbd) !== normSbd)
      );
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
                maPhach: linkedPhach || r.maPhach,
                rawScoreText: rawVal,
                score: parsedScore,
                notes: linkedPhach
                  ? `Mã phách ${linkedPhach} → SBD ${normSbd}: "${rawVal}" → ${parsedScore}`
                  : `Nhập nhanh: "${rawVal}" → ${parsedScore}`,
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
          maPhach: linkedPhach || undefined,
          score: parsedScore,
          rawScoreText: rawVal,
          scoreInWords: "",
          confidence: "CAO",
          notes: linkedPhach
            ? `Mã phách ${linkedPhach} → SBD ${normSbd}: "${rawVal}" → ${parsedScore}`
            : `Nhập nhanh: "${rawVal}" → ${parsedScore}`,
          sourceImage: "Nhập nhanh bàn phím",
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
    const parsed = parseBulkPastedScores(
      bulkPasteText,
      knownSbds,
      phachToSbdMap
    );
    if (parsed.length > 0) {
      mergeIncomingRecords(parsed, phachToSbdMap, true);
      setOcrNotice(
        `Đã dán và chuẩn hóa thành công ${parsed.length} dòng điểm.`
      );
      setShowBulkPaste(false);
      setBulkPasteText("");
    }
  };

  // ============================================================================
  // STEP 3 HANDLERS: 100% Client-Side Excel Export (Produces Hình 4)
  // ============================================================================
  const triggerBase64Download = (
    base64Data: string,
    fileName: string,
    mimeType: string
  ) => {
    const clean = base64Data.includes(",")
      ? base64Data.split(",")[1]
      : base64Data;
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
      setExcelError("Vui lòng tải lên File Excel Nguồn 1 (Bảng điểm SBD) ở Bước 1 trước.");
      setCurrentStep(1);
      return;
    }
    setIsExporting(true);
    setExcelError(null);

    try {
      const result = await exportUpdatedExcelClientSide({
        fileBase64: excelData.fileBase64,
        fileName: excelData.fileName,
        headerRow,
        sbdCol,
        scoreCol,
        sheetIndex: selectedSheetIndex,
        ocrRecords: resolvedOcrRecords.map((r) => ({
          sbd: normalizeSbdClient(r.resolvedSbd),
          maPhach: r.resolvedPhach || undefined,
          score: r.score,
        })),
        sbdToPhachMap,
        phachToSbdMap,
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

  const handleUpdateOcrRow = (
    id: string,
    field: "sbd" | "maPhach" | "rawScoreText",
    value: string
  ) => {
    setOcrRecords((prev) =>
      prev.map((rec) => {
        if (rec.id !== id) return rec;
        if (field === "maPhach") {
          const normP = normalizeSbdClient(value);
          const canonP = canonicalSbdKey(value);
          const mappedSbd =
            phachToSbdMap[normP] || phachToSbdMap[canonP] || rec.sbd;
          return {
            ...rec,
            maPhach: value,
            sbd: mappedSbd,
            notes:
              mappedSbd && mappedSbd !== normP
                ? `Khớp Mã phách ${normP} → SBD ${mappedSbd}`
                : rec.notes,
          };
        }
        if (field === "sbd") {
          const normCode = normalizeSbdClient(value);
          const canonCode = canonicalSbdKey(value);
          const mappedFromPhach =
            phachToSbdMap[normCode] || phachToSbdMap[canonCode];
          if (mappedFromPhach) {
            return {
              ...rec,
              rawSbd: value,
              maPhach: normCode,
              sbd: mappedFromPhach,
              notes: `Khớp Mã phách ${normCode} → SBD ${mappedFromPhach}`,
            };
          }
          return {
            ...rec,
            rawSbd: value,
            sbd: normCode,
            maPhach: sbdToPhachMap[normCode] || sbdToPhachMap[canonCode] || rec.maPhach,
          };
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
      knownSbds[nextIdx - 1] || `K9.${String(nextIdx).padStart(3, "0")}`;
    const defaultPhach = sbdToPhachMap[defaultSbd] || "";
    setOcrRecords((prev) => [
      ...prev,
      {
        id: `manual-${Date.now()}`,
        sbd: defaultSbd,
        rawSbd: defaultPhach || defaultSbd,
        maPhach: defaultPhach || undefined,
        score: 8.0,
        rawScoreText: "8",
        scoreInWords: "",
        confidence: "CAO",
        notes: defaultPhach
          ? `Mã phách ${defaultPhach} → SBD ${defaultSbd}`
          : "Nhập thủ công",
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
      {/* Top Bar */}
      <header className="sticky top-0 z-30 flex items-center justify-between gap-8 px-6 py-3.5 bg-white border-b border-slate-200">
        <a
          href="#top"
          className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
        >
          VietGrade OCR
        </a>

        {/* 2-Step Sequential Navigation */}
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
            <span>01. Nạp File Excel Nguồn (SBD &amp; Mã Phách)</span>
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
            <span>02. Quét PDF/Ảnh &amp; Điền Điểm Excel</span>
          </button>
        </div>

        <div className="flex items-center gap-3 shrink-0">
          <input
            ref={excel1InputRef}
            type="file"
            accept=".xlsx, .xls"
            multiple
            onChange={(e) => {
              if (e.target.files) {
                handleSmartExcelFilesUpload(e.target.files, "SLOT_1_ROSTER");
              }
            }}
            className="hidden"
          />
          <input
            ref={excel2InputRef}
            type="file"
            accept=".xlsx, .xls"
            multiple
            onChange={(e) => {
              if (e.target.files) {
                handleSmartExcelFilesUpload(e.target.files, "SLOT_2_PHACH");
              }
            }}
            className="hidden"
          />
          <button
            type="button"
            onClick={handleExecuteAndExport}
            disabled={isExporting || !excelData || ocrRecords.length === 0}
            className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 flex items-center gap-2 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            <span>
              {isExporting ? "Đang xử lý..." : "Xuất File Excel Đã Điền Điểm"}
            </span>
          </button>
        </div>
      </header>

      <main id="top" className="flex-1 max-w-[1400px] w-full mx-auto px-6 py-8 space-y-8">
        {/* Pipeline Status Strip */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 p-4 bg-white border border-slate-200 rounded-lg">
          <div className="flex flex-wrap items-center gap-6">
            <div className="flex items-center gap-3">
              <div
                className={`p-2 rounded-lg ${
                  excelData
                    ? "bg-emerald-50 text-emerald-700"
                    : "bg-slate-100 text-slate-500"
                }`}
              >
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div>
                <div className="text-[11px] text-slate-500">
                  File Excel 1 · Bảng Điểm SBD (Hình 1)
                </div>
                <div className="text-xs font-bold text-slate-900 font-mono">
                  {excelData ? excelData.fileName : "Chưa tải lên"}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3">
              <div
                className={`p-2 rounded-lg ${
                  phachExcelData
                    ? "bg-indigo-50 text-indigo-700"
                    : "bg-slate-100 text-slate-500"
                }`}
              >
                <Link2 className="w-5 h-5" />
              </div>
              <div>
                <div className="text-[11px] text-slate-500">
                  File Excel 2 · Bảng Mã Phách (Hình 2)
                </div>
                <div className="text-xs font-bold text-slate-900 font-mono">
                  {phachExcelData
                    ? `${phachExcelData.fileName} (${totalPhachMappedCount} mã phách)`
                    : "Tùy chọn (nếu chấm theo Mã Phách)"}
                </div>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-6 text-xs font-mono tabular-nums">
            <div>
              <span className="text-slate-500">Học sinh: </span>
              <span className="font-bold text-slate-900">
                {stats.totalExcelStudents}
              </span>
            </div>
            <div>
              <span className="text-slate-500">Mã phách: </span>
              <span className="font-bold text-indigo-700">
                {totalPhachMappedCount}
              </span>
            </div>
            <div>
              <span className="text-slate-500">Phiếu chấm (PDF/Ảnh): </span>
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
            BƯỚC 1: TẢI LÊN 2 FILE EXCEL NGUỒN (HÌNH 1: BẢNG ĐIỂM SBD & HÌNH 2: BẢNG MÃ PHÁCH)
           ===================================================================== */}
        {currentStep === 1 && (
          <section className="space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h1 className="text-xl font-bold text-slate-900">
                  01. Tải Lên 2 File Excel Nguồn (Bảng Điểm SBD &amp; Bảng Ghép Mã Phách)
                </h1>
                <p className="text-xs text-slate-500 mt-1">
                  Tải lên File Excel 1 (Bảng điểm có cột SBD &amp; Điểm như Hình 1) và File Excel 2 (Bảng ghép cột SBD ↔ MÃ PHÁCH như Hình 2). Bạn có thể chọn từng file hoặc bôi đen chọn cả 2 file cùng lúc.
                </p>
              </div>

              <button
                type="button"
                onClick={handleLoadSamplePhachWorkspace}
                disabled={isExcelLoading}
                className="px-3.5 py-2 text-xs font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 rounded-lg transition-colors flex items-center gap-1.5 shrink-0 cursor-pointer"
              >
                <Sparkles className="w-3.5 h-3.5" />
                <span>Thử Mẫu Hình 1 + Hình 2 + Hình 3 → Hình 4 (Anh 9)</span>
              </button>
            </div>

            {excelError && (
              <div className="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-center gap-2.5 text-xs text-red-800">
                <AlertCircle className="w-4 h-4 shrink-0 text-red-600" />
                <span>{excelError}</span>
              </div>
            )}

            {/* 2 Side-by-Side Upload Zones for File Excel 1 (Hình 1) and File Excel 2 (Hình 2) */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* Upload Zone 1: Hình 1 - Bảng Điểm Thí Sinh */}
              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsExcelDragging1(true);
                }}
                onDragLeave={() => setIsExcelDragging1(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsExcelDragging1(false);
                  if (e.dataTransfer.files) {
                    handleSmartExcelFilesUpload(
                      e.dataTransfer.files,
                      "SLOT_1_ROSTER"
                    );
                  }
                }}
                onClick={() => excel1InputRef.current?.click()}
                className={`border-2 border-dashed rounded-xl p-7 transition-colors text-center cursor-pointer flex flex-col justify-between ${
                  isExcelDragging1
                    ? "border-emerald-600 bg-emerald-50/80"
                    : excelData
                    ? "border-emerald-500 bg-white"
                    : "border-slate-300 hover:border-emerald-600 bg-white"
                }`}
              >
                <div className="space-y-2.5">
                  <div className="w-11 h-11 mx-auto rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div className="text-sm font-bold text-slate-900">
                    File Excel Nguồn 1: Bảng Điểm Thí Sinh (Hình 1)
                  </div>
                  <p className="text-xs text-slate-500 max-w-md mx-auto">
                    File chứa danh sách học sinh, cột{" "}
                    <span className="font-mono font-semibold text-slate-700">
                      SBD
                    </span>{" "}
                    (VD: K9.001) và cột{" "}
                    <span className="font-mono font-semibold text-slate-700">
                      Điểm
                    </span>{" "}
                    để xuất ra thành Hình 4.
                  </p>
                  {excelData && (
                    <div className="pt-1 text-xs font-bold text-emerald-700 font-mono">
                      ✓ Đã nạp: {excelData.fileName} ({excelData.studentRows.length} học sinh)
                    </div>
                  )}
                </div>

                <div className="mt-4">
                  <span className="px-4 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 rounded-lg hover:bg-emerald-100 transition-colors inline-flex items-center gap-1.5">
                    <Upload className="w-3.5 h-3.5" />
                    <span>
                      {excelData
                        ? "Đổi File Excel 1 (Bảng Điểm)"
                        : "Chọn File Excel 1 (.xlsx)"}
                    </span>
                  </span>
                </div>
              </div>

              {/* Upload Zone 2: Hình 2 - Bảng Ghép Mã Phách */}
              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsExcelDragging2(true);
                }}
                onDragLeave={() => setIsExcelDragging2(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsExcelDragging2(false);
                  if (e.dataTransfer.files) {
                    handleSmartExcelFilesUpload(
                      e.dataTransfer.files,
                      "SLOT_2_PHACH"
                    );
                  }
                }}
                onClick={() => excel2InputRef.current?.click()}
                className={`border-2 border-dashed rounded-xl p-7 transition-colors text-center cursor-pointer flex flex-col justify-between ${
                  isExcelDragging2
                    ? "border-indigo-600 bg-indigo-50/80"
                    : phachExcelData
                    ? "border-indigo-500 bg-white"
                    : "border-slate-300 hover:border-indigo-600 bg-white"
                }`}
              >
                <div className="space-y-2.5">
                  <div className="w-11 h-11 mx-auto rounded-lg bg-indigo-50 border border-indigo-200 flex items-center justify-center text-indigo-700">
                    <Link2 className="w-5 h-5" />
                  </div>
                  <div className="text-sm font-bold text-slate-900">
                    File Excel Nguồn 2: Bảng Ghép Mã Phách (Hình 2)
                  </div>
                  <p className="text-xs text-slate-500 max-w-md mx-auto">
                    File chứa cột{" "}
                    <span className="font-mono font-semibold text-slate-700">
                      SBD
                    </span>{" "}
                    (VD: K9.001) và cột{" "}
                    <span className="font-mono font-semibold text-slate-700">
                      MÃ PHÁCH
                    </span>{" "}
                    (VD: A40, A41, A42, A43).
                  </p>
                  {phachExcelData && (
                    <div className="pt-1 text-xs font-bold text-indigo-700 font-mono">
                      ✓ Đã nạp: {phachExcelData.fileName} (Khớp {totalPhachMappedCount} Mã phách)
                    </div>
                  )}
                </div>

                <div className="mt-4">
                  <span className="px-4 py-2 text-xs font-semibold text-indigo-800 bg-indigo-50 border border-indigo-300 rounded-lg hover:bg-indigo-100 transition-colors inline-flex items-center gap-1.5">
                    <Upload className="w-3.5 h-3.5" />
                    <span>
                      {phachExcelData
                        ? "Đổi File Excel 2 (Bảng Mã Phách)"
                        : "Chọn File Excel 2 (.xlsx)"}
                    </span>
                  </span>
                </div>
              </div>
            </div>

            {/* Configuration & Live Preview of Hình 1 + Hình 2 */}
            {(excelData || phachExcelData) && (
              <div className="space-y-5 bg-white border border-slate-200 rounded-lg p-5">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
                  <div>
                    <h2 className="text-sm font-bold text-slate-900">
                      Cấu Trúc Cột Đã Nhận Diện &amp; Đối Chiếu SBD ↔ MÃ PHÁCH
                    </h2>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Hệ thống tự động liên kết cột SBD của Hình 1 với cột SBD &amp; MÃ PHÁCH của Hình 2.
                    </p>
                  </div>

                  {excelData && (
                    <button
                      type="button"
                      onClick={() => setCurrentStep(2)}
                      className="px-4 py-2 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
                    >
                      <span>Tiếp tục: Bước 2 — Tải PDF/Ảnh Hình 3 &amp; Khớp Điểm</span>
                      <ArrowRight className="w-4 h-4" />
                    </button>
                  )}
                </div>

                {/* File Excel 1 Column Mapping */}
                {excelData && (
                  <div className="space-y-2">
                    <div className="text-xs font-bold text-emerald-900 flex items-center gap-1.5">
                      <span>1. Cấu hình File Excel 1 — Bảng Điểm Thí Sinh ({excelData.fileName}):</span>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                      <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-200">
                        <label
                          htmlFor="step1-sheet"
                          className="text-[11px] font-semibold text-slate-700 block mb-1"
                        >
                          Trang tính (Worksheet):
                        </label>
                        <select
                          id="step1-sheet"
                          value={selectedSheetIndex}
                          onChange={(e) =>
                            handleSheetChange1(Number(e.target.value))
                          }
                          className="w-full px-2 py-1 text-xs font-semibold border border-slate-300 rounded bg-white"
                        >
                          {(
                            excelData.availableSheets || [
                              {
                                index: 0,
                                name: excelData.sheetName,
                                rowCount: excelData.totalRows,
                              },
                            ]
                          ).map((s) => (
                            <option key={s.index} value={s.index}>
                              {s.name} ({s.rowCount} hàng)
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-200">
                        <label
                          htmlFor="step1-header-row"
                          className="text-[11px] font-semibold text-slate-700 block mb-1"
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
                            handleHeaderRowChange1(
                              Math.max(1, Number(e.target.value) || 6)
                            )
                          }
                          className="w-full px-2 py-1 font-mono text-xs font-semibold border border-slate-300 rounded bg-white tabular-nums"
                        />
                      </div>

                      <div className="p-2.5 bg-amber-50/60 rounded-lg border border-amber-200">
                        <label
                          htmlFor="step1-sbd-col"
                          className="text-[11px] font-semibold text-amber-900 block mb-1"
                        >
                          Cột SBD (Hình 1):
                        </label>
                        <select
                          id="step1-sbd-col"
                          value={sbdCol}
                          onChange={(e) => setSbdCol(e.target.value)}
                          className="w-full px-2 py-1 font-mono text-xs font-bold text-amber-900 border border-amber-300 rounded bg-white"
                        >
                          {excelData.headers.map((h) => (
                            <option key={h.colLetter} value={h.colLetter}>
                              Cột {h.colLetter} ({h.label})
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="p-2.5 bg-blue-50/60 rounded-lg border border-blue-200">
                        <label
                          htmlFor="step1-score-col"
                          className="text-[11px] font-semibold text-blue-900 block mb-1"
                        >
                          Cột Điểm Cần Điền (Hình 4):
                        </label>
                        <select
                          id="step1-score-col"
                          value={scoreCol}
                          onChange={(e) => setScoreCol(e.target.value)}
                          className="w-full px-2 py-1 font-mono text-xs font-bold text-blue-700 border border-blue-300 rounded bg-white"
                        >
                          {excelData.headers.map((h) => (
                            <option key={h.colLetter} value={h.colLetter}>
                              Cột {h.colLetter} ({h.label})
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>
                )}

                {/* File Excel 2 Column Mapping (Hình 2: SBD <-> MÃ PHÁCH) */}
                {phachExcelData && (
                  <div className="space-y-2 pt-2 border-t border-slate-200">
                    <div className="text-xs font-bold text-indigo-900 flex items-center justify-between">
                      <span>
                        2. Cấu hình File Excel 2 — Bảng Ghép Mã Phách ({phachExcelData.fileName}):
                      </span>
                      <button
                        type="button"
                        onClick={() => setPhachExcelData(null)}
                        className="text-[11px] font-normal text-slate-500 hover:text-red-600 cursor-pointer"
                      >
                        Gỡ File Excel 2
                      </button>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                      <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-200">
                        <label
                          htmlFor="step1-phach-sheet"
                          className="text-[11px] font-semibold text-slate-700 block mb-1"
                        >
                          Trang tính (Worksheet):
                        </label>
                        <select
                          id="step1-phach-sheet"
                          value={phachSheetIndex}
                          onChange={(e) =>
                            handleSheetChange2(Number(e.target.value))
                          }
                          className="w-full px-2 py-1 text-xs font-semibold border border-slate-300 rounded bg-white"
                        >
                          {(
                            phachExcelData.availableSheets || [
                              {
                                index: 0,
                                name: phachExcelData.sheetName,
                                rowCount: phachExcelData.totalRows,
                              },
                            ]
                          ).map((s) => (
                            <option key={s.index} value={s.index}>
                              {s.name} ({s.rowCount} hàng)
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-200">
                        <label
                          htmlFor="step1-phach-header-row"
                          className="text-[11px] font-semibold text-slate-700 block mb-1"
                        >
                          Hàng tiêu đề (Header Row):
                        </label>
                        <input
                          id="step1-phach-header-row"
                          type="number"
                          min={1}
                          max={20}
                          value={phachHeaderRow}
                          onChange={(e) =>
                            handleHeaderRowChange2(
                              Math.max(1, Number(e.target.value) || 4)
                            )
                          }
                          className="w-full px-2 py-1 font-mono text-xs font-semibold border border-slate-300 rounded bg-white tabular-nums"
                        />
                      </div>

                      <div className="p-2.5 bg-amber-50/60 rounded-lg border border-amber-200">
                        <label
                          htmlFor="step1-phach-sbd-col"
                          className="text-[11px] font-semibold text-amber-900 block mb-1"
                        >
                          Cột SBD (Hình 2):
                        </label>
                        <select
                          id="step1-phach-sbd-col"
                          value={phachSbdCol}
                          onChange={(e) => setPhachSbdCol(e.target.value)}
                          className="w-full px-2 py-1 font-mono text-xs font-bold text-amber-900 border border-amber-300 rounded bg-white"
                        >
                          {phachExcelData.headers.map((h) => (
                            <option key={h.colLetter} value={h.colLetter}>
                              Cột {h.colLetter} ({h.label})
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="p-2.5 bg-indigo-50/70 rounded-lg border border-indigo-200">
                        <label
                          htmlFor="step1-phach-code-col"
                          className="text-[11px] font-semibold text-indigo-900 block mb-1"
                        >
                          Cột MÃ PHÁCH (Hình 2):
                        </label>
                        <select
                          id="step1-phach-code-col"
                          value={phachCodeCol}
                          onChange={(e) => setPhachCodeCol(e.target.value)}
                          className="w-full px-2 py-1 font-mono text-xs font-bold text-indigo-700 border border-indigo-300 rounded bg-white"
                        >
                          {phachExcelData.headers.map((h) => (
                            <option key={h.colLetter} value={h.colLetter}>
                              Cột {h.colLetter} ({h.label})
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>
                )}

                {/* Combined Roster + Mã Phách Preview */}
                {excelData && (
                  <div className="border border-slate-200 rounded-lg overflow-hidden">
                    <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-700 flex justify-between items-center">
                      <span>
                        Xem trước danh sách học sinh &amp; Mã Phách đã khớp (Từ hàng{" "}
                        {headerRow + 1} trở đi):
                      </span>
                      <span className="font-mono tabular-nums text-slate-500">
                        Tổng cộng: {excelData.studentRows.length} học sinh
                      </span>
                    </div>
                    <div className="overflow-x-auto max-h-[280px]">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 font-semibold text-slate-600">
                          <tr>
                            <th className="py-2 px-3 w-14 text-center">Hàng</th>
                            {excelData.headers.map((h) => (
                              <React.Fragment key={h.colLetter}>
                                <th className="py-2 px-3 whitespace-nowrap">
                                  {h.label} ({h.colLetter})
                                </th>
                                {h.colLetter === sbdCol && phachExcelData && (
                                  <th className="py-2 px-3 whitespace-nowrap bg-indigo-50 text-indigo-900 border-x border-indigo-200">
                                    MÃ PHÁCH (Khớp Hình 2)
                                  </th>
                                )}
                              </React.Fragment>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {matchedStudentRows.slice(0, 15).map((row) => (
                            <tr key={row.rowNumber} className="hover:bg-slate-50">
                              <td className="py-2 px-3 text-center font-mono tabular-nums text-slate-400">
                                {row.rowNumber}
                              </td>
                              {excelData.headers.map((h) => (
                                <React.Fragment key={h.colLetter}>
                                  <td
                                    className={`py-2 px-3 ${
                                      h.colLetter === sbdCol
                                        ? "font-mono font-bold text-slate-900 bg-amber-50/50 tabular-nums"
                                        : "text-slate-700"
                                    }`}
                                  >
                                    {String(row.cells[h.colLetter] ?? "")}
                                  </td>
                                  {h.colLetter === sbdCol && phachExcelData && (
                                    <td className="py-2 px-3 font-mono font-bold text-indigo-700 bg-indigo-50/40 border-x border-indigo-100 tabular-nums">
                                      {row.linkedPhach ? (
                                        row.linkedPhach
                                      ) : (
                                        <span className="text-slate-400 font-normal italic">
                                          Trống
                                        </span>
                                      )}
                                    </td>
                                  )}
                                </React.Fragment>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}
          </section>
        )}

        {/* =====================================================================
            BƯỚC 2: TẢI FILE PDF / ẢNH PHIẾU CHẤM (HÌNH 3) & XUẤT EXCEL (HÌNH 4)
           ===================================================================== */}
        {currentStep === 2 && (
          <section className="space-y-5">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
              <div>
                <h1 className="text-xl font-bold text-slate-900">
                  02. Tải File PDF / Ảnh Phiếu Chấm (Hình 3) &amp; Xuất File Excel (Hình 4)
                </h1>
                <p className="text-xs text-slate-500 mt-0.5">
                  Hệ thống tự động đọc cột{" "}
                  <span className="font-mono font-semibold text-slate-700">
                    MÃ PHÁCH
                  </span>{" "}
                  (hoặc SBD) và{" "}
                  <span className="font-mono font-semibold text-slate-700">
                    ĐIỂM BẰNG SỐ
                  </span>{" "}
                  trên phiếu chấm (Hình 3), đối chiếu qua Bảng Mã Phách (Hình 2) để điền điểm vào đúng học sinh trên Bảng Điểm (Hình 4).
                </p>
              </div>

              <input
                ref={imageInputRef}
                type="file"
                accept="image/*,.pdf,application/pdf"
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
                    Dán Nhanh Danh Sách Điểm (Theo Mã Phách hoặc SBD)
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
                  placeholder={"A40  8\nA41  8,50\nA42  8\nA43  8"}
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
              {/* Left Column (5 cols): Grading Sheet Photo / PDF Page Viewer (Hình 3) */}
              <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg flex flex-col overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-2 bg-slate-50">
                  <div className="flex items-center gap-2 text-xs font-medium text-slate-700 truncate">
                    <span>Phiếu Chấm ({images.length})</span>
                    {activeImage && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="font-mono text-slate-500 truncate">
                          {activeImage.name}
                        </span>
                      </>
                    )}
                  </div>

                  <div className="flex items-center gap-1.5 shrink-0">
                    <button
                      type="button"
                      onClick={() => imageInputRef.current?.click()}
                      className="px-2.5 py-1 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded transition-colors cursor-pointer flex items-center gap-1"
                      title="Tải thêm file Ảnh hoặc PDF nhiều trang"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Tải Ảnh / PDF</span>
                    </button>
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
                      <div className="text-xs text-slate-500 max-w-xs mx-auto">
                        Tải lên file PDF (gộp nhiều trang phiếu chấm) hoặc ảnh chụp Phiếu Chấm Bài Thi (Hình 3).
                      </div>
                      <button
                        type="button"
                        onClick={() => imageInputRef.current?.click()}
                        className="px-4 py-2 text-xs font-bold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
                      >
                        Chọn file PDF hoặc Ảnh phiếu chấm
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Right Column (7 cols): Extracted OCR Table (MÃ PHÁCH <-> SBD <-> ĐIỂM) */}
              <div className="lg:col-span-7 bg-white border border-slate-200 rounded-lg flex flex-col overflow-hidden">
                {engineMode === "RAPID_KEYBOARD" && excelData ? (
                  <>
                    <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-4 bg-blue-50/60">
                      <div className="text-xs font-semibold text-blue-950">
                        Nhập Nhanh Theo Danh Sách (Gõ{" "}
                        <code className="font-mono">8,5</code> + nhấn{" "}
                        <code className="font-mono">Enter</code>)
                      </div>
                      <div className="flex items-center gap-3">
                        <span className="text-xs font-mono tabular-nums text-blue-800 font-bold">
                          Đã nhập: {stats.matchedCount}/{excelData.studentRows.length}
                        </span>
                        <button
                          type="button"
                          onClick={() => setEngineMode("HYBRID_CLOUD")}
                          className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer"
                        >
                          Về Bảng OCR
                        </button>
                      </div>
                    </div>

                    <div className="flex-1 overflow-auto max-h-[540px]">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 text-[11px] font-semibold text-slate-600">
                          <tr>
                            <th className="py-2.5 px-3 w-12 text-center">Hàng</th>
                            <th className="py-2.5 px-3">SBD (Cột {sbdCol})</th>
                            {phachExcelData && (
                              <th className="py-2.5 px-3 text-indigo-800">
                                Mã Phách (Hình 2)
                              </th>
                            )}
                            <th className="py-2.5 px-3">Họ và tên</th>
                            <th className="py-2.5 px-3 w-32">Gõ Điểm</th>
                            <th className="py-2.5 px-3 text-right w-24">
                              Điểm (Cột {scoreCol})
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {matchedStudentRows.map((row, rIdx) => {
                            const norm = row.normalizedSbd;
                            const inputVal =
                              rapidInputs[norm] ??
                              (row.matchedRawText !== null
                                ? row.matchedRawText
                                : "");
                            const studentName =
                              row.cells["B"] ||
                              row.cells["C"] ||
                              Object.values(row.cells).find(
                                (v) =>
                                  typeof v === "string" && v !== row.displaySbd
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
                                {phachExcelData && (
                                  <td className="py-2 px-3 font-mono font-bold text-indigo-700 tabular-nums">
                                    {row.linkedPhach || (
                                      <span className="text-slate-300 font-normal">
                                        Trống
                                      </span>
                                    )}
                                  </td>
                                )}
                                <td className="py-2 px-3 text-slate-700 truncate max-w-[180px]">
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
                                      handleRapidScoreChange(
                                        row.displaySbd,
                                        row.linkedPhach,
                                        e.target.value
                                      )
                                    }
                                    onKeyDown={(e) =>
                                      handleRapidKeyDown(
                                        e,
                                        rIdx,
                                        matchedStudentRows.length
                                      )
                                    }
                                    className="w-24 px-2.5 py-1 font-mono font-semibold text-slate-900 bg-white border border-slate-300 rounded focus:outline-none focus:border-blue-600 focus:ring-2 focus:ring-blue-100 tabular-nums"
                                  />
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold tabular-nums text-sm">
                                  {row.matchedScore !== null ? (
                                    <span className="text-emerald-700">
                                      {row.matchedScore}
                                    </span>
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
                        <span>
                          Bảng Kết Quả OCR &amp; Khớp Phách (Mã Phách ↔ SBD ↔ Điểm)
                        </span>
                        <span aria-hidden="true">·</span>
                        <span className="font-mono tabular-nums text-slate-500">
                          {resolvedOcrRecords.length} bản ghi
                        </span>
                      </div>

                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setEngineMode("RAPID_KEYBOARD")}
                          className="px-2.5 py-1 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded transition-colors cursor-pointer"
                        >
                          Chuyển sang Gõ Nhanh
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
                              <span>
                                {ocrProgressText ||
                                  "Đang xử lý hàng loạt trang phiếu chấm..."}
                              </span>
                            </div>
                            <span className="font-mono tabular-nums">
                              {ocrProgressPct}%
                            </span>
                          </div>
                          <div className="w-full h-1.5 bg-blue-100 rounded-full overflow-hidden">
                            <div
                              className="h-full bg-blue-600 transition-all duration-200"
                              style={{ width: `${ocrProgressPct}%` }}
                            />
                          </div>
                        </div>
                      )}

                      {isOcrRunning && resolvedOcrRecords.length === 0 ? (
                        <div className="p-6 space-y-2.5">
                          {[1, 2, 3, 4, 5].map((n) => (
                            <div
                              key={n}
                              className="h-9 bg-slate-100 animate-pulse rounded border border-slate-200"
                            />
                          ))}
                        </div>
                      ) : resolvedOcrRecords.length === 0 ? (
                        <div className="p-12 text-center space-y-4">
                          <div className="text-sm font-semibold text-slate-700">
                            Chưa có dữ liệu điểm nào từ Phiếu Chấm (Hình 3)
                          </div>
                          <div className="flex items-center justify-center gap-3">
                            <button
                              type="button"
                              onClick={() => imageInputRef.current?.click()}
                              className="px-3.5 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg cursor-pointer"
                            >
                              Tải File PDF / Ảnh Phiếu Chấm
                            </button>
                            <button
                              type="button"
                              onClick={() => setEngineMode("RAPID_KEYBOARD")}
                              className="px-3.5 py-2 text-xs font-semibold text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-lg cursor-pointer"
                            >
                              Gõ Nhanh Theo SBD
                            </button>
                          </div>
                        </div>
                      ) : (
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="border-b border-slate-200 bg-slate-50/80 text-[11px] font-semibold text-slate-600">
                              <th className="py-2.5 px-3 w-10 text-center">#</th>
                              <th className="py-2.5 px-3">
                                Mã Phách (Hình 3 &amp; 2)
                              </th>
                              <th className="py-2.5 px-3">
                                SBD Khớp (Hình 1 &amp; 2)
                              </th>
                              <th className="py-2.5 px-3">
                                Điểm Gốc (Hình 3)
                              </th>
                              <th className="py-2.5 px-3 text-right">
                                Điểm Chuẩn (Hình 4)
                              </th>
                              <th className="py-2.5 px-3 w-10 text-center">
                                Xóa
                              </th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200 text-xs">
                            {resolvedOcrRecords.map((rec, index) => {
                              const isMatchedInExcel =
                                knownSbds.includes(
                                  normalizeSbdClient(rec.resolvedSbd)
                                ) ||
                                knownSbds.some(
                                  (k) =>
                                    canonicalSbdKey(k) ===
                                    canonicalSbdKey(rec.resolvedSbd)
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
                                    <input
                                      type="text"
                                      placeholder="—"
                                      value={rec.resolvedPhach || ""}
                                      onChange={(e) =>
                                        handleUpdateOcrRow(
                                          rec.id,
                                          "maPhach",
                                          e.target.value
                                        )
                                      }
                                      className="w-20 px-2 py-1 font-mono font-bold text-indigo-700 bg-indigo-50/50 border border-indigo-200 rounded focus:outline-none focus:border-indigo-600 focus:bg-white tabular-nums"
                                    />
                                  </td>
                                  <td className="py-2 px-3">
                                    <div className="flex items-center gap-2">
                                      <input
                                        type="text"
                                        value={rec.resolvedSbd}
                                        onChange={(e) =>
                                          handleUpdateOcrRow(
                                            rec.id,
                                            "sbd",
                                            e.target.value
                                          )
                                        }
                                        className="w-24 px-2 py-1 font-mono font-semibold text-slate-900 bg-slate-50 border border-slate-300 rounded focus:outline-none focus:border-blue-600 focus:bg-white tabular-nums"
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

            {/* Live Matched Excel Preview (Hình 4: Kết Quả So Khớp SBD Hình 1 × Mã Phách Hình 2 × Điểm Hình 3) */}
            {excelData && resolvedOcrRecords.length > 0 && (
              <div className="bg-white border-2 border-emerald-600/80 rounded-lg overflow-hidden shadow-xs">
                <div className="px-4 py-3 bg-emerald-50/80 border-b border-emerald-200 flex items-center justify-between">
                  <div className="text-xs font-bold text-emerald-950 flex items-center gap-2">
                    <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                    <span>
                      XEM TRƯỚC FILE EXCEL XUẤT RA (HÌNH 4): Đã khớp{" "}
                      {stats.matchedCount}/{stats.totalExcelStudents} học sinh
                      {phachExcelData
                        ? ` (Qua cầu nối SBD ↔ MÃ PHÁCH)`
                        : ""}
                    </span>
                  </div>
                </div>

                {/* Excel Sheet Title Header Reproduction (Rows 1-4) */}
                <div className="px-5 py-3 bg-white border-b border-slate-200 text-xs font-serif text-slate-900">
                  {excelData.titleRows && excelData.titleRows.length > 0 ? (
                    <div className="space-y-0.5">
                      {excelData.titleRows.map((tr) => (
                        <div
                          key={tr.rowNumber}
                          className={
                            tr.rowNumber >= 3
                              ? "text-center font-bold text-sm mt-1"
                              : "text-[13px] font-semibold tracking-tight"
                          }
                        >
                          {tr.texts.join(" — ")}
                        </div>
                      ))}
                    </div>
                  ) : (
                    <>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-1">
                        <div>
                          <div className="text-[13px] tracking-tight">
                            SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH NINH BÌNH
                          </div>
                          <div className="text-[13px] font-bold pl-8">
                            TRƯỜNG THCS MỸ XÁ
                          </div>
                        </div>
                      </div>
                      <div className="text-center font-bold text-sm mt-1">
                        BẢNG ĐIỂM THI
                      </div>
                    </>
                  )}
                </div>

                {/* Matched Excel Table (Exact Hình 4 Layout) */}
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
                                {val !== null && val !== undefined
                                  ? String(val)
                                  : ""}
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
            {excelData && resolvedOcrRecords.length > 0 && (
              <div className="max-w-2xl mx-auto w-full bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-5">
                <div className="flex items-center justify-between">
                  <div>
                    <h2 className="text-base font-bold text-slate-900">
                      Xuất File Excel Hoàn Chỉnh (Hình 4)
                    </h2>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Đã khớp SBD ↔ Mã Phách và điền điểm số thực Float vào cột{" "}
                      {scoreCol}, bảo toàn 100% định dạng File Excel 1.
                    </p>
                  </div>
                  <div className="w-10 h-10 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700">
                    <FileCheck2 className="w-5 h-5" />
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-3 text-xs border-t border-b border-slate-200 py-3.5 bg-slate-50/60 rounded-lg px-3">
                  <div className="text-center">
                    <div className="text-slate-500 text-[11px]">
                      Tiêu đề Sở/Trường
                    </div>
                    <div className="font-mono text-emerald-700 font-bold mt-0.5">
                      Bảo toàn ({excelData?.mergedCellsCount || 0} Merges)
                    </div>
                  </div>
                  <div className="text-center border-x border-slate-200">
                    <div className="text-slate-500 text-[11px]">
                      Khớp SBD ↔ Mã Phách
                    </div>
                    <div className="font-mono text-indigo-700 font-bold mt-0.5">
                      {stats.matchedCount}/{stats.totalExcelStudents} học sinh
                    </div>
                  </div>
                  <div className="text-center">
                    <div className="text-slate-500 text-[11px]">
                      Kiểu dữ liệu điểm
                    </div>
                    <div className="font-mono text-blue-700 font-bold mt-0.5">
                      Số thực (Float)
                    </div>
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
                  disabled={
                    isExporting || !excelData || resolvedOcrRecords.length === 0
                  }
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
