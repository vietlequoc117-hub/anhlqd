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
  Sparkles,
} from "lucide-react";
import {
  OcrRecord,
  ExcelSheetData,
  normalizeSbdClient,
  parseVietnameseScoreClient,
  generateSampleGradingSheetImage,
} from "./utils/sampleGradingSheet";

interface UploadedImage {
  id: string;
  name: string;
  mimeType: string;
  dataUrl: string;
}

export default function App() {
  // Step Navigation State (Step 1 -> Step 2 -> Step 3)
  const [currentStep, setCurrentStep] = useState<1 | 2 | 3>(1);

  // STEP 1: Excel Source File State
  const [excelData, setExcelData] = useState<ExcelSheetData | null>(null);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [headerRow, setHeaderRow] = useState<number>(5);
  const [sbdCol, setSbdCol] = useState<string>("G");
  const [scoreCol, setScoreCol] = useState<string>("I");
  const [isExcelLoading, setIsExcelLoading] = useState<boolean>(false);
  const [excelError, setExcelError] = useState<string | null>(null);
  const [isExcelDragging, setIsExcelDragging] = useState<boolean>(false);

  // STEP 2: Grading Sheet Images & OCR State
  const [images, setImages] = useState<UploadedImage[]>([]);
  const [activeImageIdx, setActiveImageIdx] = useState<number>(0);
  const [imageZoom, setImageZoom] = useState<number>(1);
  const [isOcrRunning, setIsOcrRunning] = useState<boolean>(false);
  const [ocrError, setOcrError] = useState<string | null>(null);
  const [ocrRecords, setOcrRecords] = useState<OcrRecord[]>([]);

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
    # BƯỚC 1: Đã có file Excel nguồn "${fileName}"
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    sheet_names = wb.sheetnames
    ws = wb[sheet_names[${selectedSheetIndex}]] if len(sheet_names) > ${selectedSheetIndex} else wb.active

    HEADER_ROW = ${headerRow}
    START_DATA_ROW = ${startRow}
    SBD_COL = "${sbdCol}"
    SCORE_COL = "${scoreCol}"

    # BƯỚC 2: Dữ liệu (SBD -> Điểm số Float) đã trích xuất OCR từ ảnh PHIẾU CHẤM BÀI THI
    ocr_scores: dict[str, float] = {
${dictLines || "    # Dữ liệu từ các ảnh phiếu chấm tải lên"}
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

  // Parse Excel file buffer
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

      const response = await fetch("/api/excel/parse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileBase64: base64,
          fileName,
          sheetIndex,
        }),
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Lỗi khi phân tích cấu trúc file Excel.");
      }

      setExcelData({
        ...data,
        fileBase64: base64,
      });
      setSelectedSheetIndex(sheetIndex);
      setHeaderRow(data.detectedHeaderRow || 5);
      setSbdCol(data.detectedSbdCol || "G");
      setScoreCol(data.detectedScoreCol || "I");
      setExportResult(null);

      // Automatically guide user forward to Step 2
      setCurrentStep(2);
    } catch (err: any) {
      setExcelError(err?.message || "Không thể đọc file Excel. Vui lòng đảm bảo file định dạng .xlsx hợp lệ.");
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
      setExcelError("Vui lòng tải lên file Excel có định dạng .xlsx hoặc .xls.");
      return;
    }
    await parseAndLoadExcelBuffer(file, file.name, 0);
  };

  const handleSheetChange = async (newSheetIdx: number) => {
    if (!excelData) return;
    setSelectedSheetIndex(newSheetIdx);
    setIsExcelLoading(true);
    try {
      const response = await fetch("/api/excel/parse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileBase64: excelData.fileBase64,
          fileName: excelData.fileName,
          sheetIndex: newSheetIdx,
        }),
      });
      const data = await response.json();
      if (response.ok) {
        setExcelData({
          ...data,
          fileBase64: excelData.fileBase64,
        });
        setHeaderRow(data.detectedHeaderRow || 5);
        setSbdCol(data.detectedSbdCol || "G");
        setScoreCol(data.detectedScoreCol || "I");
      }
    } catch {
      // Keep existing sheet
    } finally {
      setIsExcelLoading(false);
    }
  };

  const handleHeaderRowChange = async (newHeaderRow: number) => {
    setHeaderRow(newHeaderRow);
    if (!excelData) return;
    try {
      const response = await fetch("/api/excel/parse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileBase64: excelData.fileBase64,
          fileName: excelData.fileName,
          headerRow: newHeaderRow,
          sbdCol,
          scoreCol,
          sheetIndex: selectedSheetIndex,
        }),
      });
      const data = await response.json();
      if (response.ok) {
        setExcelData({
          ...data,
          fileBase64: excelData.fileBase64,
        });
      }
    } catch {
      // Keep existing
    }
  };

  // Helper to load sample files on demand (not on page load)
  const handleLoadSampleWorkspace = async () => {
    setIsExcelLoading(true);
    setOcrError(null);
    setExcelError(null);

    try {
      const excelRes = await fetch("/api/sample-excel");
      const excelJson = await excelRes.json();
      if (!excelRes.ok) {
        throw new Error(excelJson.error || "Không thể tải file Excel mẫu.");
      }
      setExcelData(excelJson);
      setSelectedSheetIndex(0);
      setHeaderRow(excelJson.detectedHeaderRow || 5);
      setSbdCol(excelJson.detectedSbdCol || "G");
      setScoreCol(excelJson.detectedScoreCol || "I");

      const sampleImg = generateSampleGradingSheetImage();
      const sampleImageObj: UploadedImage = {
        id: `sample-${Date.now()}`,
        name: sampleImg.name,
        mimeType: sampleImg.mimeType,
        dataUrl: sampleImg.dataUrl,
      };
      setImages([sampleImageObj]);
      setActiveImageIdx(0);

      await runGeminiOcrOnImages([sampleImageObj], false);
      setCurrentStep(2);
    } catch (err: any) {
      setExcelError(err?.message || "Lỗi khi khởi tạo dữ liệu mẫu.");
    } finally {
      setIsExcelLoading(false);
    }
  };

  // Step 2 OCR execution
  const runGeminiOcrOnImages = async (targetImages: UploadedImage[], append = false) => {
    if (targetImages.length === 0) return;
    setIsOcrRunning(true);
    setOcrError(null);

    try {
      const response = await fetch("/api/ocr", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          images: targetImages.map((img) => ({
            name: img.name,
            mimeType: img.mimeType,
            data: img.dataUrl,
          })),
        }),
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Không thể trích xuất dữ liệu từ ảnh.");
      }

      const incoming: OcrRecord[] = data.records || [];
      setOcrRecords((prev) => {
        const combined = append ? [...prev, ...incoming] : incoming;
        const bySbd = new Map<string, OcrRecord>();
        for (const item of combined) {
          bySbd.set(normalizeSbdClient(item.sbd), item);
        }
        return Array.from(bySbd.values());
      });
      setExportResult(null);
    } catch (err: any) {
      setOcrError(err?.message || "Lỗi kết nối máy chủ OCR.");
    } finally {
      setIsOcrRunning(false);
    }
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    setOcrError(null);

    const newImages: UploadedImage[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      newImages.push({
        id: `${file.name}-${Date.now()}-${i}`,
        name: file.name,
        mimeType: file.type || "image/png",
        dataUrl,
      });
    }

    const updatedList = [...images, ...newImages];
    setImages(updatedList);
    setActiveImageIdx(updatedList.length - newImages.length);
    await runGeminiOcrOnImages(newImages, true);
    if (imageInputRef.current) imageInputRef.current.value = "";
  };

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
      const response = await fetch("/api/excel/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
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
        }),
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Lỗi khi ghi điểm vào file Excel.");
      }

      setExportResult(data);
      triggerBase64Download(
        data.updatedExcelBase64,
        data.outputFileName,
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
    const defaultSbd = `K6.${String(nextIdx).padStart(3, "0")}`;
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
      {/* Top Bar: Brand, Step Tabs, Primary Action */}
      <header className="sticky top-0 z-30 flex items-center justify-between gap-8 px-6 py-3.5 bg-white border-b border-slate-200">
        <a
          href="#top"
          className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
        >
          VietGrade OCR
        </a>

        {/* 3-Step Wizard Navigation */}
        <div className="flex items-center gap-1.5 p-1 bg-slate-100 rounded-lg">
          <button
            type="button"
            onClick={() => setCurrentStep(1)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors flex items-center gap-2 cursor-pointer ${
              currentStep === 1
                ? "bg-white text-slate-900 shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <span className={`w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-mono ${
              excelData ? "bg-emerald-600 text-white" : "bg-slate-300 text-slate-700"
            }`}>
              {excelData ? "✓" : "1"}
            </span>
            <span>01. Nạp File Excel Nguồn</span>
          </button>

          <button
            type="button"
            onClick={() => setCurrentStep(2)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors flex items-center gap-2 cursor-pointer ${
              currentStep === 2
                ? "bg-white text-slate-900 shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <span className={`w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-mono ${
              ocrRecords.length > 0 ? "bg-blue-600 text-white" : "bg-slate-300 text-slate-700"
            }`}>
              {ocrRecords.length > 0 ? "✓" : "2"}
            </span>
            <span>02. Tải Ảnh Phiếu Chấm (OCR)</span>
          </button>

          <button
            type="button"
            onClick={() => setCurrentStep(3)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors flex items-center gap-2 cursor-pointer ${
              currentStep === 3
                ? "bg-white text-slate-900 shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <span className="w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-mono bg-slate-300 text-slate-700">
              3
            </span>
            <span>03. Khớp &amp; Xuất File Excel</span>
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
            className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 flex items-center gap-2 cursor-pointer shadow-xs"
          >
            <Download className="w-3.5 h-3.5" />
            <span>{isExporting ? "Đang xử lý..." : "Xuất File Excel Đã Điền Điểm"}</span>
          </button>
        </div>
      </header>

      <main id="top" className="flex-1 max-w-[1400px] w-full mx-auto px-6 py-8 space-y-8">
        {/* Pipeline Summary Bar */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 p-4 bg-white border border-slate-200 rounded-lg">
          <div className="flex items-center gap-4">
            <div className={`p-2.5 rounded-lg ${excelData ? "bg-emerald-50 text-emerald-700" : "bg-slate-100 text-slate-500"}`}>
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <div className="text-xs font-medium text-slate-500">File Excel Nguồn (Bước 1):</div>
              <div className="text-sm font-bold text-slate-900 font-mono">
                {excelData ? excelData.fileName : "Chưa tải lên file Excel nguồn"}
              </div>
            </div>
          </div>

          <div className="flex items-center gap-6 text-xs font-mono tabular-nums">
            <div>
              <span className="text-slate-500">Học sinh trong file: </span>
              <span className="font-bold text-slate-900">{stats.totalExcelStudents}</span>
            </div>
            <div>
              <span className="text-slate-500">Ảnh phiếu chấm: </span>
              <span className="font-bold text-slate-900">{images.length}</span>
            </div>
            <div>
              <span className="text-slate-500">Đã khớp điểm: </span>
              <span className="font-bold text-emerald-700">{stats.matchedCount}</span>
            </div>
          </div>
        </div>

        {/* =====================================================================
            BƯỚC 1: TẢI LÊN FILE EXCEL NGUỒN TRƯỚC TIÊN
           ===================================================================== */}
        {currentStep === 1 && (
          <section className="space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 text-xs font-bold text-white bg-emerald-700 rounded">
                    BƯỚC 1
                  </span>
                  <h2 className="text-xl font-bold text-slate-900">
                    Tải Lên File Excel Nguồn Cần Nhập Điểm
                  </h2>
                </div>
                <p className="text-xs text-slate-600 mt-1">
                  Hãy chọn hoặc kéo thả file bảng điểm học sinh của bạn vào đây. Hệ thống sẽ đọc danh sách học sinh,
                  xác định cột <strong>Số BD</strong> (ví dụ cột G) và cột <strong>Điểm</strong> (ví dụ cột I).
                </p>
              </div>

              <div className="flex items-center gap-3 shrink-0">
                <button
                  type="button"
                  onClick={handleLoadSampleWorkspace}
                  disabled={isExcelLoading}
                  className="px-3.5 py-2 text-xs font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isExcelLoading ? "animate-spin" : ""}`} />
                  <span>Dùng Thử File Mẫu</span>
                </button>

                <button
                  type="button"
                  onClick={() => excelInputRef.current?.click()}
                  className="px-4 py-2 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  <FileUp className="w-4 h-4" />
                  <span>Chọn File Excel Từ Máy Tính</span>
                </button>
              </div>
            </div>

            {excelError && (
              <div className="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-center gap-2.5 text-xs text-red-800">
                <AlertCircle className="w-4 h-4 shrink-0 text-red-600" />
                <span>{excelError}</span>
              </div>
            )}

            {/* Dropzone */}
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setIsExcelDragging(true);
              }}
              onDragLeave={() => setIsExcelDragging(false)}
              onDrop={handleExcelDrop}
              onClick={() => excelInputRef.current?.click()}
              className={`border-2 border-dashed rounded-xl p-10 transition-all text-center cursor-pointer ${
                isExcelDragging
                  ? "border-emerald-600 bg-emerald-50/80 scale-[1.005]"
                  : "border-slate-300 hover:border-emerald-500 bg-white"
              }`}
            >
              <div className="max-w-md mx-auto flex flex-col items-center space-y-3">
                <div className="w-14 h-14 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 shadow-xs">
                  <FileSpreadsheet className="w-7 h-7" />
                </div>
                <div className="text-base font-bold text-slate-900">
                  {excelData ? (
                    <span>
                      Đã nạp file: <span className="text-emerald-700 font-mono">{excelData.fileName}</span>
                    </span>
                  ) : (
                    <span>Kéo &amp; thả file Excel nguồn (.xlsx) vào đây hoặc bấm để duyệt file</span>
                  )}
                </div>
                <p className="text-xs text-slate-500 leading-relaxed">
                  Bảo toàn 100% dòng tiêu đề trường/sở (hàng 1-4), merged cells, độ rộng cột, font chữ và viền kẻ bảng.
                </p>
                <div className="pt-2">
                  <span className="px-4 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-300 rounded-lg hover:bg-emerald-100 transition-colors inline-flex items-center gap-1.5">
                    <Upload className="w-3.5 h-3.5" />
                    <span>{excelData ? "Đổi sang file Excel khác" : "Chọn file Excel từ máy"}</span>
                  </span>
                </div>
              </div>
            </div>

            {/* If Excel Loaded: Show Structure Preview & Column Settings */}
            {excelData && (
              <div className="space-y-4 bg-white border border-slate-200 rounded-lg p-5">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
                  <div>
                    <h3 className="text-sm font-bold text-slate-900">
                      Cấu Trúc Bảng Điểm Đã Nhận Diện Thành Công
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Kiểm tra hàng tiêu đề, cột Số BD và cột Điểm. Bạn có thể điều chỉnh lại nếu cần.
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={() => setCurrentStep(2)}
                    className="px-4 py-2 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer shadow-xs"
                  >
                    <span>Tiếp tục: Bước 2 - Tải Ảnh Phiếu Chấm</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                    <label htmlFor="step1-header-row" className="text-xs font-semibold text-slate-700 block mb-1">
                      Hàng tiêu đề:
                    </label>
                    <input
                      id="step1-header-row"
                      type="number"
                      min={1}
                      max={20}
                      value={headerRow}
                      onChange={(e) => handleHeaderRowChange(Math.max(1, Number(e.target.value) || 5))}
                      className="w-full px-2.5 py-1.5 font-mono text-sm font-semibold border border-slate-300 rounded bg-white tabular-nums"
                    />
                    <span className="text-[11px] text-slate-500 mt-1 block">
                      Dữ liệu học sinh sẽ được đọc từ hàng {headerRow + 1} trở đi.
                    </span>
                  </div>

                  <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                    <label htmlFor="step1-sbd-col" className="text-xs font-semibold text-slate-700 block mb-1">
                      Cột Số Báo Danh (SBD):
                    </label>
                    <select
                      id="step1-sbd-col"
                      value={sbdCol}
                      onChange={(e) => setSbdCol(e.target.value)}
                      className="w-full px-2.5 py-1.5 font-mono text-sm font-semibold border border-slate-300 rounded bg-white"
                    >
                      {excelData.headers.map((h) => (
                        <option key={h.colLetter} value={h.colLetter}>
                          Cột {h.colLetter} ({h.label})
                        </option>
                      ))}
                    </select>
                    <span className="text-[11px] text-slate-500 mt-1 block">
                      Cột chứa mã Số Báo Danh để đối chiếu VLOOKUP.
                    </span>
                  </div>

                  <div className="p-3 bg-blue-50/60 rounded-lg border border-blue-200">
                    <label htmlFor="step1-score-col" className="text-xs font-semibold text-blue-900 block mb-1">
                      Cột Điểm (Chỉ ghi đè cột này):
                    </label>
                    <select
                      id="step1-score-col"
                      value={scoreCol}
                      onChange={(e) => setScoreCol(e.target.value)}
                      className="w-full px-2.5 py-1.5 font-mono text-sm font-bold text-blue-700 border border-blue-300 rounded bg-white"
                    >
                      {excelData.headers.map((h) => (
                        <option key={h.colLetter} value={h.colLetter}>
                          Cột {h.colLetter} ({h.label})
                        </option>
                      ))}
                    </select>
                    <span className="text-[11px] text-blue-800 mt-1 block">
                      Điểm thi (Float) sẽ chỉ được ghi vào cột này.
                    </span>
                  </div>
                </div>

                {/* Quick Student Roster Preview Table */}
                <div className="mt-4 border border-slate-200 rounded-lg overflow-hidden">
                  <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-700 flex justify-between items-center">
                    <span>Xem trước {Math.min(excelData.studentRows.length, 6)} học sinh đầu tiên:</span>
                    <span className="font-mono text-slate-500">Tổng cộng: {excelData.studentRows.length} học sinh</span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100 border-b border-slate-200 font-semibold text-slate-600">
                          <th className="py-2 px-3 w-12 text-center">Hàng</th>
                          {excelData.headers.slice(0, 7).map((h) => (
                            <th key={h.colLetter} className="py-2 px-3">
                              {h.label} ({h.colLetter})
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {excelData.studentRows.slice(0, 6).map((row) => (
                          <tr key={row.rowNumber} className="hover:bg-slate-50">
                            <td className="py-2 px-3 text-center font-mono text-slate-400">{row.rowNumber}</td>
                            {excelData.headers.slice(0, 7).map((h) => (
                              <td key={h.colLetter} className={`py-2 px-3 ${h.colLetter === sbdCol ? "font-mono font-bold text-slate-900 bg-amber-50/50" : "text-slate-700"}`}>
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
            BƯỚC 2: TẢI ẢNH PHIẾU CHẤM BÀI THI & CHẠY OCR CHỮ VIẾT TAY
           ===================================================================== */}
        {currentStep === 2 && (
          <section className="space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 text-xs font-bold text-white bg-blue-700 rounded">
                    BƯỚC 2
                  </span>
                  <h2 className="text-xl font-bold text-slate-900">
                    Tải Lên Các File Ảnh Chứa Điểm (Phiếu Chấm Bài Thi)
                  </h2>
                </div>
                <p className="text-xs text-slate-600 mt-1">
                  File nguồn đã nạp: <strong className="font-mono text-slate-900">{excelData?.fileName || "Chưa tải"}</strong>.
                  Bây giờ hãy tải các ảnh chụp phiếu chấm bài thi viết tay để AI tự động trích xuất mã SBD và Điểm số.
                </p>
              </div>

              <div className="flex items-center gap-3 shrink-0">
                <button
                  type="button"
                  onClick={() => setCurrentStep(1)}
                  className="px-3 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors flex items-center gap-1 cursor-pointer"
                >
                  <ArrowLeft className="w-3.5 h-3.5" />
                  <span>Quay lại Bước 1</span>
                </button>

                <input
                  ref={imageInputRef}
                  type="file"
                  accept="image/*"
                  multiple
                  onChange={handleImageUpload}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={() => imageInputRef.current?.click()}
                  className="px-4 py-2 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  <Upload className="w-4 h-4" />
                  <span>Chọn Ảnh Phiếu Chấm</span>
                </button>

                <button
                  type="button"
                  onClick={() => setCurrentStep(3)}
                  disabled={ocrRecords.length === 0}
                  className="px-4 py-2 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 rounded-lg transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <span>Sang Bước 3: Khớp Điểm</span>
                  <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </div>

            {ocrError && (
              <div className="p-3.5 bg-red-50 border border-red-200 rounded-lg flex items-center gap-2.5 text-xs text-red-800">
                <AlertCircle className="w-4 h-4 shrink-0 text-red-600" />
                <span>{ocrError}</span>
              </div>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
              {/* Image Viewer */}
              <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg flex flex-col overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-2 bg-slate-50">
                  <div className="flex items-center gap-2 text-xs font-medium text-slate-700 truncate">
                    <span>Ảnh Phiếu Chấm ({images.length})</span>
                    {activeImage && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="font-mono text-slate-500 truncate">{activeImage.name}</span>
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
                  <div className="px-4 py-2 border-b border-slate-200 flex items-center gap-1.5 overflow-x-auto bg-white">
                    {images.map((img, idx) => (
                      <button
                        key={img.id}
                        type="button"
                        onClick={() => setActiveImageIdx(idx)}
                        className={`px-2.5 py-1 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          idx === activeImageIdx
                            ? "bg-slate-900 text-white"
                            : "bg-slate-100 text-slate-600 hover:text-slate-900"
                        }`}
                      >
                        Phiếu #{idx + 1}: {img.name}
                      </button>
                    ))}
                  </div>
                )}

                <div className="flex-1 min-h-[420px] max-h-[540px] overflow-auto bg-slate-100 p-4 flex items-start justify-center">
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
                    <div className="my-auto text-center p-8 space-y-3">
                      <ScanText className="w-8 h-8 text-slate-400 mx-auto" />
                      <div className="text-sm font-semibold text-slate-700">
                        Chưa có ảnh Phiếu Chấm Bài Thi nào được chọn
                      </div>
                      <p className="text-xs text-slate-500 max-w-xs mx-auto">
                        Tải lên một hoặc nhiều ảnh chụp bảng điểm viết tay để AI tự động nhận diện chữ số.
                      </p>
                      <button
                        type="button"
                        onClick={() => imageInputRef.current?.click()}
                        className="px-4 py-2 text-xs font-bold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
                      >
                        Chọn ảnh phiếu chấm từ máy tính
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* Extracted OCR Table */}
              <div className="lg:col-span-7 bg-white border border-slate-200 rounded-lg flex flex-col overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-4 bg-slate-50">
                  <div className="flex items-center gap-2 text-xs font-medium text-slate-700">
                    <span>Kết Quả OCR (SBD — Điểm Số)</span>
                    <span aria-hidden="true">·</span>
                    <span className="font-mono tabular-nums text-slate-500">
                      {ocrRecords.length} bản ghi
                    </span>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleAddOcrRow}
                      className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 flex items-center gap-1 cursor-pointer"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Thêm dòng</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => runGeminiOcrOnImages(images, false)}
                      disabled={isOcrRunning || images.length === 0}
                      className="px-2.5 py-1.5 text-xs font-semibold text-slate-800 bg-slate-100 hover:bg-slate-200 rounded transition-colors whitespace-nowrap shrink-0 flex items-center gap-1 cursor-pointer"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isOcrRunning ? "animate-spin" : ""}`} />
                      <span>Quét lại</span>
                    </button>
                  </div>
                </div>

                <div className="flex-1 overflow-auto max-h-[540px]">
                  {isOcrRunning ? (
                    <div className="p-8 space-y-3">
                      <div className="flex items-center gap-2 text-xs font-medium text-blue-700">
                        <RefreshCw className="w-4 h-4 animate-spin" />
                        <span>
                          AI đang phân tích nét số viết tay (1 vs 7 gạch ngang, 4 hở/kín, dấu phẩy thập phân)...
                        </span>
                      </div>
                      {[1, 2, 3, 4, 5, 6].map((n) => (
                        <div
                          key={n}
                          className="h-9 bg-slate-100 animate-pulse rounded border border-slate-200"
                        />
                      ))}
                    </div>
                  ) : ocrRecords.length === 0 ? (
                    <div className="p-12 text-center space-y-2">
                      <div className="text-sm font-medium text-slate-700">
                        Chưa có dữ liệu SBD &amp; Điểm từ ảnh
                      </div>
                      <p className="text-xs text-slate-500">
                        Nhấn nút "Chọn Ảnh Phiếu Chấm" để bắt đầu nhận diện.
                      </p>
                    </div>
                  ) : (
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-slate-200 bg-slate-50/80 text-[11px] font-semibold text-slate-600">
                          <th className="py-2.5 px-3 w-12 text-center">#</th>
                          <th className="py-2.5 px-3">SBD Chuẩn Hóa</th>
                          <th className="py-2.5 px-3">Điểm Viết Tay (Gốc)</th>
                          <th className="py-2.5 px-3 text-right">Điểm Chuẩn (Float)</th>
                          <th className="py-2.5 px-3">Đối Chiếu Điểm Chữ</th>
                          <th className="py-2.5 px-3 w-12 text-center">Xóa</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 text-xs">
                        {ocrRecords.map((rec, index) => {
                          const isMatchedInExcel = excelData
                            ? excelData.studentRows.some(
                                (r) =>
                                  normalizeSbdClient(r.cells[sbdCol] ?? r.sbd) ===
                                  normalizeSbdClient(rec.sbd)
                              )
                            : false;

                          return (
                            <tr key={rec.id} className="hover:bg-slate-50/80 transition-colors">
                              <td className="py-2 px-3 text-center font-mono tabular-nums text-slate-400">
                                {index + 1}
                              </td>
                              <td className="py-2 px-3">
                                <div className="flex items-center gap-2">
                                  <input
                                    type="text"
                                    value={rec.sbd}
                                    onChange={(e) => handleUpdateOcrRow(rec.id, "sbd", e.target.value)}
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
                                    handleUpdateOcrRow(rec.id, "rawScoreText", e.target.value)
                                  }
                                  className="w-20 px-2 py-1 font-mono text-slate-700 bg-slate-50 border border-slate-300 rounded focus:outline-none focus:border-blue-600 focus:bg-white tabular-nums"
                                />
                              </td>
                              <td className="py-2 px-3 text-right font-mono font-bold text-blue-700 tabular-nums text-sm">
                                {rec.score}
                              </td>
                              <td className="py-2 px-3 text-slate-600 truncate max-w-[200px]">
                                {rec.scoreInWords ? `"${rec.scoreInWords}"` : rec.notes}
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
              </div>
            </div>
          </section>
        )}

        {/* =====================================================================
            BƯỚC 3: ĐỐI CHIẾU KHỚP VLOOKUP, XUẤT EXCEL & MÃ PYTHON OPENPYXL
           ===================================================================== */}
        {currentStep === 3 && (
          <section className="space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 text-xs font-bold text-white bg-blue-900 rounded">
                    BƯỚC 3
                  </span>
                  <h2 className="text-xl font-bold text-slate-900">
                    Đối Chiếu Khớp Điểm &amp; Xuất File Excel (Bảo Toàn Định Dạng)
                  </h2>
                </div>
                <p className="text-xs text-slate-600 mt-1">
                  Đã khớp <strong className="text-emerald-700 font-mono">{stats.matchedCount}</strong> điểm học sinh.
                  Chỉ ghi đè giá trị Float vào cột <strong>Điểm (Cột {scoreCol})</strong>, không làm vỡ các vùng gộp ô,
                  tiêu đề hay viền kẻ bảng của file nguồn.
                </p>
              </div>

              <div className="flex items-center gap-3 shrink-0">
                <button
                  type="button"
                  onClick={() => setCurrentStep(2)}
                  className="px-3 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors flex items-center gap-1 cursor-pointer"
                >
                  <ArrowLeft className="w-3.5 h-3.5" />
                  <span>Quay lại Bước 2</span>
                </button>

                <button
                  type="button"
                  onClick={handleExecuteAndExport}
                  disabled={isExporting || !excelData || ocrRecords.length === 0}
                  className="px-4 py-2 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  <Download className="w-4 h-4" />
                  <span>{isExporting ? "Đang xử lý..." : "Tải File Excel Đã Điền Điểm"}</span>
                </button>
              </div>
            </div>

            {/* Excel Preview Grid with Match Status */}
            <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
              <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="flex items-center gap-2 text-xs text-slate-600">
                  <FileSpreadsheet className="w-4 h-4 text-emerald-700 shrink-0" />
                  <span className="font-bold text-slate-900">{excelData?.fileName}</span>
                  <span aria-hidden="true">·</span>
                  <span>Sheet: {excelData?.sheetName}</span>
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  <div className="flex items-center gap-1 p-1 bg-slate-200/70 rounded-lg">
                    <button
                      type="button"
                      onClick={() => setRowFilter("ALL")}
                      className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors cursor-pointer ${
                        rowFilter === "ALL" ? "bg-white text-slate-900 shadow-xs" : "text-slate-600"
                      }`}
                    >
                      Tất cả ({stats.totalExcelStudents})
                    </button>
                    <button
                      type="button"
                      onClick={() => setRowFilter("MATCHED")}
                      className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors cursor-pointer ${
                        rowFilter === "MATCHED" ? "bg-white text-emerald-800 shadow-xs" : "text-slate-600"
                      }`}
                    >
                      Đã khớp ({stats.matchedCount})
                    </button>
                    <button
                      type="button"
                      onClick={() => setRowFilter("UNMATCHED")}
                      className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors cursor-pointer ${
                        rowFilter === "UNMATCHED" ? "bg-white text-amber-800 shadow-xs" : "text-slate-600"
                      }`}
                    >
                      Giữ trống ({stats.unmatchedExcelCount})
                    </button>
                  </div>

                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      placeholder="Tìm tên hoặc SBD..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="pl-8 pr-3 py-1 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-blue-600 w-40"
                    />
                  </div>
                </div>
              </div>

              <div className="overflow-x-auto max-h-[440px]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="sticky top-0 z-10 bg-slate-100 border-b border-slate-300">
                    <tr className="text-[11px] font-mono text-slate-500 bg-slate-200/60 border-b border-slate-300">
                      <th className="py-1 px-2.5 text-center border-r border-slate-300 w-14">Hàng</th>
                      {excelData?.headers.map((h) => (
                        <th
                          key={h.colLetter}
                          className={`py-1 px-3 text-center border-r border-slate-300 ${
                            h.colLetter === sbdCol
                              ? "bg-amber-100 text-amber-900 font-bold"
                              : h.colLetter === scoreCol
                              ? "bg-blue-100 text-blue-900 font-bold"
                              : ""
                          }`}
                        >
                          {h.colLetter} {h.colLetter === sbdCol ? "(Số BD)" : h.colLetter === scoreCol ? "(Điểm - Ghi đè)" : ""}
                        </th>
                      ))}
                      <th className="py-1 px-3 text-left">Trạng thái VLOOKUP</th>
                    </tr>

                    <tr className="text-xs font-semibold text-slate-800 bg-slate-100">
                      <th className="py-2 px-2.5 text-center font-mono text-slate-500 border-r border-slate-300">{headerRow}</th>
                      {excelData?.headers.map((h) => (
                        <th key={h.colLetter} className={`py-2 px-3 border-r border-slate-200 whitespace-nowrap ${h.colLetter === scoreCol ? "bg-blue-50 text-blue-900 font-bold" : ""}`}>
                          {h.label}
                        </th>
                      ))}
                      <th className="py-2 px-3 whitespace-nowrap text-slate-600">Đối chiếu OCR → Excel</th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-200">
                    {filteredStudentRows.map((row) => (
                      <tr key={row.rowNumber} className={`hover:bg-slate-50 transition-colors ${row.isMatched ? "" : "bg-slate-50/40"}`}>
                        <td className="py-2 px-2.5 text-center font-mono tabular-nums text-slate-400 bg-slate-50 border-r border-slate-200">
                          {row.rowNumber}
                        </td>
                        {excelData?.headers.map((h) => {
                          const isScore = h.colLetter === scoreCol;
                          const isSbd = h.colLetter === sbdCol;
                          const rawCellVal = row.cells[h.colLetter];

                          if (isScore) {
                            return (
                              <td
                                key={h.colLetter}
                                className={`py-2 px-3 border-r border-slate-200 font-mono tabular-nums text-right ${
                                  row.isMatched ? "bg-blue-50/60 font-bold text-blue-700" : "text-slate-400 italic"
                                }`}
                              >
                                {row.isMatched ? (
                                  <div className="flex items-center justify-end gap-1.5">
                                    <span className="text-[11px] font-normal text-slate-400">{scoreCol}{row.rowNumber}:</span>
                                    <span className="text-sm">{row.matchedScore}</span>
                                  </div>
                                ) : (
                                  <span className="text-[11px] not-italic text-slate-400">(Trống)</span>
                                )}
                              </td>
                            );
                          }

                          return (
                            <td key={h.colLetter} className={`py-2 px-3 border-r border-slate-200 ${isSbd ? "font-mono font-bold text-slate-900 bg-amber-50/40" : "text-slate-700"}`}>
                              {rawCellVal !== null && rawCellVal !== undefined ? String(rawCellVal) : ""}
                            </td>
                          );
                        })}

                        <td className="py-2 px-3 whitespace-nowrap">
                          {row.isMatched ? (
                            <div className="flex items-center gap-1.5 text-xs text-emerald-700 font-medium">
                              <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
                              <span>Khớp SBD {row.normalizedSbd}</span>
                              <span aria-hidden="true">·</span>
                              <span className="font-mono tabular-nums text-slate-600">
                                "{row.matchedRawText}" → {row.matchedScore}
                              </span>
                            </div>
                          ) : (
                            <div className="flex items-center gap-1.5 text-xs text-slate-500">
                              <span>Chưa có điểm</span>
                              <span aria-hidden="true">·</span>
                              <span>Giữ ô trống</span>
                            </div>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Python Code & Export Card */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
              <div className="lg:col-span-7 bg-[#0F172A] text-slate-100 rounded-lg border border-slate-800 flex flex-col overflow-hidden">
                <div className="px-4 py-2.5 border-b border-slate-800 flex items-center justify-between bg-slate-900/80">
                  <div className="flex items-center gap-2 text-xs text-slate-300 font-mono">
                    <Code2 className="w-3.5 h-3.5 text-blue-400" />
                    <span>auto_update_scores_openpyxl.py</span>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleCopyCode}
                      className="px-2.5 py-1 text-xs font-medium text-slate-200 bg-slate-800 hover:bg-slate-700 rounded transition-colors flex items-center gap-1.5 cursor-pointer"
                    >
                      {copiedCode ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                      <span>{copiedCode ? "Đã chép" : "Sao chép"}</span>
                    </button>
                    <button
                      type="button"
                      onClick={handleDownloadPythonScript}
                      className="px-2.5 py-1 text-xs font-medium text-slate-200 bg-slate-800 hover:bg-slate-700 rounded transition-colors cursor-pointer"
                    >
                      Tải file .py
                    </button>
                  </div>
                </div>

                <pre className="p-4 text-xs font-mono leading-relaxed overflow-x-auto flex-1 max-h-[360px] text-slate-200">
                  <code>{liveOpenpyxlScript}</code>
                </pre>
              </div>

              <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg p-5 flex flex-col justify-between space-y-4">
                <div className="space-y-3">
                  <div className="flex items-start justify-between">
                    <div>
                      <h3 className="text-sm font-bold text-slate-900">
                        Kiểm Định Bảo Toàn Định Dạng Excel
                      </h3>
                      <p className="text-xs text-slate-500 mt-0.5">
                        Chỉ ghi đè giá trị Float vào cột {scoreCol}, bảo toàn 100% định dạng.
                      </p>
                    </div>
                    <FileCheck2 className="w-5 h-5 text-blue-700 shrink-0" />
                  </div>

                  <div className="space-y-2 text-xs border-t border-b border-slate-200 py-3">
                    <div className="flex justify-between">
                      <span className="text-slate-600">Tiêu đề Sở/Trường:</span>
                      <span className="font-mono text-emerald-700 font-medium">Bảo toàn ({excelData?.mergedCellsCount || 0} Merges)</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-slate-600">Font chữ &amp; Khung viền:</span>
                      <span className="font-mono text-slate-900">Giữ nguyên 100%</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-slate-600">Kiểu dữ liệu điểm:</span>
                      <span className="font-mono font-bold text-blue-700">Float/Number</span>
                    </div>
                  </div>

                  {exportResult && (
                    <pre className="p-3 bg-slate-900 text-emerald-300 font-mono text-[11px] rounded border border-slate-800 overflow-x-auto whitespace-pre-wrap leading-relaxed">
                      {exportResult.pythonAuditLog}
                    </pre>
                  )}
                </div>

                <button
                  type="button"
                  onClick={handleExecuteAndExport}
                  disabled={isExporting || !excelData || ocrRecords.length === 0}
                  className="w-full px-4 py-2.5 text-xs font-bold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer shadow-sm"
                >
                  <Download className="w-4 h-4" />
                  <span>{exportResult ? `Tải Lại ${exportResult.outputFileName}` : "Xuất & Tải File Excel Đã Nhập Điểm"}</span>
                </button>
              </div>
            </div>
          </section>
        )}
      </main>

      <footer className="mt-12 border-t border-slate-200 bg-white py-4 px-6">
        <div className="max-w-[1400px] mx-auto flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500">
          <span>Quy trình 3 bước chuẩn: Tải file Excel nguồn → Tải ảnh phiếu chấm → Khớp VLOOKUP &amp; Xuất file</span>
          <span>Bảo toàn 100% định dạng openpyxl · Số thực Float/Number</span>
        </div>
      </footer>
    </div>
  );
}
