import {
  ExcelSheetData,
  OcrRecord,
  normalizeSbdClient,
  canonicalSbdKey,
} from "./sampleGradingSheet";

export interface SavedExcelWorkspace {
  id: string;
  fileName: string;
  subjectName: string;
  excelData: ExcelSheetData;
  selectedSheetIndex: number;
  headerRow: number;
  sbdCol: string;
  scoreCol: string;
  ocrRecords: OcrRecord[];
  scannedImageNames: string[];
  totalStudents: number;
  matchedCount: number;
  createdAt: number;
  updatedAt: number;
}

const DB_NAME = "vietgrade_ocr_storage_v1";
const STORE_NAME = "excel_workspaces";
const ACTIVE_ID_KEY = "vietgrade_active_workspace_id";

function openWorkspaceDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB not supported"));
      return;
    }
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export function buildWorkspaceId(fileName: string, sheetIndex = 0): string {
  const cleanName = fileName.trim().toLowerCase().replace(/\s+/g, "_");
  return `ws__${cleanName}__sheet_${sheetIndex}`;
}

export function inferSubjectFromExcel(
  fileName: string,
  titleRows?: { rowNumber: number; texts: string[] }[]
): string {
  const combinedText = [
    fileName.replace(/\.xlsx?$/i, ""),
    ...(titleRows ? titleRows.flatMap((r) => r.texts) : []),
  ]
    .join(" ")
    .toLowerCase();

  const subjectPatterns: { pattern: RegExp; label: string }[] = [
    { pattern: /\b(ngữ\s*văn|văn|van)\b/i, label: "Ngữ Văn" },
    { pattern: /\b(toán|toan)\b/i, label: "Toán" },
    { pattern: /\b(tiếng\s*anh|anh\s*văn|ngoại\s*ngữ|anh)\b/i, label: "Tiếng Anh" },
    { pattern: /\b(vật\s*lý|vật\s*lí|lý|li)\b/i, label: "Vật Lý" },
    { pattern: /\b(hóa\s*học|hóa|hoa)\b/i, label: "Hóa Học" },
    { pattern: /\b(sinh\s*học|sinh)\b/i, label: "Sinh Học" },
    { pattern: /\b(lịch\s*sử|sử|su)\b/i, label: "Lịch Sử" },
    { pattern: /\b(địa\s*lý|địa\s*lí|địa|dia)\b/i, label: "Địa Lý" },
    { pattern: /\b(gdcd|giáo\s*dục\s*công\s*dân|ktpl)\b/i, label: "GDCD / KTPL" },
    { pattern: /\b(tin\s*học|tin)\b/i, label: "Tin Học" },
    { pattern: /\b(công\s*nghệ|cn)\b/i, label: "Công Nghệ" },
    { pattern: /\b(khtn|khoa\s*học\s*tự\s*nhiên)\b/i, label: "KHTN" },
  ];

  for (const item of subjectPatterns) {
    if (item.pattern.test(combinedText)) {
      return item.label;
    }
  }

  const baseName = fileName.replace(/\.xlsx?$/i, "").trim();
  return baseName || "Môn học";
}

export function computeWorkspaceMatchStats(
  excelData: ExcelSheetData,
  sbdCol: string,
  ocrRecords: OcrRecord[]
): {
  totalStudents: number;
  matchedCount: number;
  unmatchedCount: number;
  unmatchedSbds: string[];
} {
  const lookup = new Map<string, OcrRecord>();
  for (const rec of ocrRecords) {
    const norm = normalizeSbdClient(rec.sbd);
    const canon = canonicalSbdKey(rec.sbd);
    if (norm) lookup.set(norm, rec);
    if (canon) lookup.set(canon, rec);
  }

  let matchedCount = 0;
  const unmatchedSbds: string[] = [];

  for (const row of excelData.studentRows) {
    const rawSbd = row.cells[sbdCol] ?? row.sbd;
    const normSbd = normalizeSbdClient(rawSbd);
    const canonSbd = canonicalSbdKey(rawSbd);
    const isMatched = Boolean(
      (normSbd && lookup.get(normSbd)) || (canonSbd && lookup.get(canonSbd))
    );
    if (isMatched) {
      matchedCount++;
    } else if (rawSbd !== null && rawSbd !== undefined && String(rawSbd).trim()) {
      unmatchedSbds.push(String(rawSbd).trim());
    }
  }

  const totalStudents = excelData.studentRows.length;
  return {
    totalStudents,
    matchedCount,
    unmatchedCount: Math.max(0, totalStudents - matchedCount),
    unmatchedSbds,
  };
}

export async function loadAllWorkspaces(): Promise<SavedExcelWorkspace[]> {
  try {
    const db = await openWorkspaceDb();
    return await new Promise<SavedExcelWorkspace[]>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readonly");
      const store = tx.objectStore(STORE_NAME);
      const req = store.getAll();
      req.onsuccess = () => {
        const list = (req.result as SavedExcelWorkspace[]) || [];
        list.sort((a, b) => b.updatedAt - a.updatedAt);
        resolve(list);
      };
      req.onerror = () => reject(req.error);
    });
  } catch {
    return [];
  }
}

export async function saveWorkspace(workspace: SavedExcelWorkspace): Promise<void> {
  try {
    const db = await openWorkspaceDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readwrite");
      const store = tx.objectStore(STORE_NAME);
      store.put(workspace);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // Ignore storage errors if private mode blocks IndexedDB
  }
}

export async function deleteWorkspaceFromStorage(id: string): Promise<void> {
  try {
    const db = await openWorkspaceDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readwrite");
      const store = tx.objectStore(STORE_NAME);
      store.delete(id);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // Ignore storage errors
  }
}

export function getActiveWorkspaceId(): string | null {
  try {
    return localStorage.getItem(ACTIVE_ID_KEY);
  } catch {
    return null;
  }
}

export function setActiveWorkspaceId(id: string | null): void {
  try {
    if (id) {
      localStorage.setItem(ACTIVE_ID_KEY, id);
    } else {
      localStorage.removeItem(ACTIVE_ID_KEY);
    }
  } catch {
    // Ignore
  }
}
