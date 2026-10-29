import * as pdfjsLib from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export interface ExtractedPdfPageImage {
  name: string;
  mimeType: string;
  dataUrl: string;
}

/**
 * Converts a multi-page scanned PDF file into high-resolution compressed JPEG images (one per page)
 * 100% client-side in the browser so every page can be previewed and processed by OCR sequentially.
 */
export async function convertPdfFileToImages(
  file: File,
  onProgress?: (currentPage: number, totalPages: number) => void
): Promise<ExtractedPdfPageImage[]> {
  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(arrayBuffer),
  });

  const pdfDoc = await loadingTask.promise;
  const numPages = pdfDoc.numPages;
  const baseName = file.name.replace(/\.pdf$/i, "");
  const pageImages: ExtractedPdfPageImage[] = [];

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    onProgress?.(pageNum, numPages);
    const page = await pdfDoc.getPage(pageNum);

    // Compute scale so the longest side is ~1600px (sharp enough for handwriting, compact for fast OCR)
    const unscaledViewport = page.getViewport({ scale: 1 });
    const maxDim = Math.max(unscaledViewport.width, unscaledViewport.height, 1);
    const targetScale = Math.min(2.5, Math.max(1.5, 1600 / maxDim));
    const viewport = page.getViewport({ scale: targetScale });

    const canvas = document.createElement("canvas");
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    const ctx = canvas.getContext("2d");

    if (!ctx) continue;

    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    await page.render({
      canvasContext: ctx,
      viewport,
      canvas,
    } as any).promise;

    const dataUrl = canvas.toDataURL("image/jpeg", 0.88);
    pageImages.push({
      name:
        numPages > 1
          ? `${baseName} - Trang ${pageNum}/${numPages}`
          : `${baseName} (PDF)`,
      mimeType: "image/jpeg",
      dataUrl,
    });
  }

  return pageImages;
}
