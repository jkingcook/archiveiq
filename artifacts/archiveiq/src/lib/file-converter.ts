import { createRequire } from "module";
import { logger } from "./logger.js";

const _require = createRequire(import.meta.url);

export type ConvertedFile =
  | { mode: "image"; base64: string; mediaType: "image/jpeg" | "image/png" | "image/webp" | "image/gif"; filename: string; pages?: number }
  | { mode: "pdf-doc"; base64: string; filename: string; pages?: number }
  | { mode: "text"; text: string; filename: string }
  | { mode: "multi-text"; texts: string[]; filename: string }
  | { mode: "multi-image"; images: Array<{ base64: string; mediaType: string; page: number }>; filename: string; pages: number };

const NATIVE_IMAGE_EXTS = new Set(["jpg", "jpeg", "png", "webp", "gif"]);
const CONVERTIBLE_IMAGE_EXTS = new Set(["bmp", "tiff", "tif", "heic", "heif", "avif", "svg", "ico", "jfif"]);
const TEXT_EXTS = new Set(["txt", "md", "rtf", "csv", "tsv", "json", "xml", "html", "htm", "log", "nfo"]);

function ext(filename: string): string {
  return filename.toLowerCase().split(".").pop() ?? "";
}

function mediaTypeForExt(e: string): "image/jpeg" | "image/png" | "image/webp" | "image/gif" {
  if (e === "png") return "image/png";
  if (e === "webp") return "image/webp";
  if (e === "gif") return "image/gif";
  return "image/jpeg";
}

async function trySharpConvert(buffer: Buffer, filename: string): Promise<Buffer | null> {
  try {
    const sharp = _require("sharp");
    return await sharp(buffer).png().toBuffer();
  } catch (e) {
    logger.warn({ filename, err: String(e) }, "sharp conversion failed");
    return null;
  }
}

async function tryPdfText(buffer: Buffer): Promise<{ text: string; pages: number } | null> {
  try {
    const pdfParse = _require("pdf-parse");
    const data = await pdfParse(buffer);
    if (data.text && data.text.trim().length > 100) {
      return { text: data.text, pages: data.numpages ?? 1 };
    }
    return null;
  } catch (e) {
    logger.warn({ err: String(e) }, "pdf-parse failed");
    return null;
  }
}

async function tryMammoth(buffer: Buffer): Promise<string | null> {
  try {
    const mammoth = _require("mammoth");
    const result = await mammoth.extractRawText({ buffer });
    return result.value ?? null;
  } catch (e) {
    logger.warn({ err: String(e) }, "mammoth extraction failed");
    return null;
  }
}

async function tryXlsx(buffer: Buffer): Promise<string | null> {
  try {
    const XLSX = _require("xlsx");
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const lines: string[] = [];
    for (const sheetName of workbook.SheetNames as string[]) {
      lines.push(`=== Sheet: ${sheetName} ===`);
      const sheet = workbook.Sheets[sheetName];
      lines.push(XLSX.utils.sheet_to_csv(sheet));
    }
    return lines.join("\n");
  } catch (e) {
    logger.warn({ err: String(e) }, "xlsx extraction failed");
    return null;
  }
}

async function detectMimeByMagic(buffer: Buffer): Promise<string | null> {
  try {
    const { fileTypeFromBuffer } = await import("file-type");
    const result = await fileTypeFromBuffer(buffer);
    return result?.mime ?? null;
  } catch (e) {
    logger.warn({ err: String(e) }, "file-type detection failed");
    return null;
  }
}

export async function convertFile(buffer: Buffer, filename: string): Promise<ConvertedFile> {
  const e = ext(filename);
  logger.info({ filename, ext: e, size: buffer.length }, "[converter] processing file");

  // ── Native Claude image types ──────────────────────────────────────
  if (NATIVE_IMAGE_EXTS.has(e)) {
    logger.info({ filename }, "[converter] native image → direct");
    return { mode: "image", base64: buffer.toString("base64"), mediaType: mediaTypeForExt(e), filename };
  }

  // ── Convertible image types (via sharp) ────────────────────────────
  if (CONVERTIBLE_IMAGE_EXTS.has(e)) {
    logger.info({ filename }, "[converter] non-standard image → sharp");
    const converted = await trySharpConvert(buffer, filename);
    if (converted) {
      return { mode: "image", base64: converted.toString("base64"), mediaType: "image/png", filename };
    }
    // Fallback: send raw and let Claude try
    return { mode: "image", base64: buffer.toString("base64"), mediaType: "image/jpeg", filename };
  }

  // ── PDF ────────────────────────────────────────────────────────────
  if (e === "pdf") {
    logger.info({ filename }, "[converter] PDF → trying text extraction");
    const textResult = await tryPdfText(buffer);
    if (textResult && textResult.text.trim().length > 100) {
      logger.info({ filename, pages: textResult.pages, chars: textResult.text.length }, "[converter] PDF text extracted");
      return { mode: "text", text: `[PDF — ${textResult.pages} page(s)]\n\n${textResult.text.substring(0, 80000)}`, filename };
    }
    // Scanned PDF → use Claude's native PDF document type
    logger.info({ filename }, "[converter] PDF → Claude native document type (scanned/image-based)");
    return { mode: "pdf-doc", base64: buffer.toString("base64"), filename };
  }

  // ── DOCX ───────────────────────────────────────────────────────────
  if (e === "docx") {
    logger.info({ filename }, "[converter] DOCX → mammoth");
    const text = await tryMammoth(buffer);
    if (text && text.trim().length > 0) {
      return { mode: "text", text: `[DOCX document]\n\n${text.substring(0, 80000)}`, filename };
    }
    return { mode: "text", text: `[DOCX file: ${filename} — text extraction failed; binary document]`, filename };
  }

  // ── Excel ──────────────────────────────────────────────────────────
  if (e === "xlsx" || e === "xls") {
    logger.info({ filename }, "[converter] Excel → xlsx");
    const text = await tryXlsx(buffer);
    if (text && text.trim().length > 0) {
      return { mode: "text", text: `[Excel spreadsheet]\n\n${text.substring(0, 80000)}`, filename };
    }
    return { mode: "text", text: `[Excel file: ${filename} — extraction failed]`, filename };
  }

  // ── CSV / TSV ──────────────────────────────────────────────────────
  if (e === "csv" || e === "tsv") {
    const text = buffer.toString("utf8", 0, Math.min(buffer.length, 80000));
    return { mode: "text", text: `[${e.toUpperCase()} data]\n\n${text}`, filename };
  }

  // ── PPTX ──────────────────────────────────────────────────────────
  if (e === "pptx" || e === "ppt") {
    // Try xlsx library (it can partially read PPTX)
    const text = await tryXlsx(buffer);
    if (text && text.trim().length > 10) {
      return { mode: "text", text: `[PowerPoint presentation — extracted text]\n\n${text.substring(0, 80000)}`, filename };
    }
    return { mode: "text", text: `[PowerPoint file: ${filename} — could not extract text automatically. Please describe the content.]`, filename };
  }

  // ── Text formats ───────────────────────────────────────────────────
  if (TEXT_EXTS.has(e)) {
    logger.info({ filename }, "[converter] plain text file");
    return { mode: "text", text: buffer.toString("utf8", 0, Math.min(buffer.length, 100000)), filename };
  }

  // ── ZIP / TAR (extract and process first image/text found) ─────────
  if (e === "zip") {
    logger.info({ filename }, "[converter] ZIP → note (extraction not yet supported)");
    return { mode: "text", text: `[ZIP archive: ${filename}, ${buffer.length} bytes. Contains compressed files — please describe what you know about this archive's contents.]`, filename };
  }

  // ── EML / MSG ─────────────────────────────────────────────────────
  if (e === "eml" || e === "msg") {
    const text = buffer.toString("utf8", 0, Math.min(buffer.length, 80000));
    const printable = text.replace(/[^\x20-\x7E\n\r\t]/g, " ");
    return { mode: "text", text: `[Email file: ${filename}]\n\n${printable}`, filename };
  }

  // ── Unknown: detect by magic bytes ────────────────────────────────
  logger.info({ filename }, "[converter] unknown extension → magic byte detection");
  const mime = await detectMimeByMagic(buffer);
  if (mime) {
    logger.info({ filename, mime }, "[converter] detected MIME type");
    if (mime === "application/pdf") {
      return { mode: "pdf-doc", base64: buffer.toString("base64"), filename };
    }
    if (mime.startsWith("image/")) {
      const knownMime = ["image/jpeg", "image/png", "image/webp", "image/gif"].includes(mime)
        ? (mime as "image/jpeg" | "image/png" | "image/webp" | "image/gif")
        : null;
      if (knownMime) {
        return { mode: "image", base64: buffer.toString("base64"), mediaType: knownMime, filename };
      }
      const converted = await trySharpConvert(buffer, filename);
      if (converted) {
        return { mode: "image", base64: converted.toString("base64"), mediaType: "image/png", filename };
      }
      return { mode: "image", base64: buffer.toString("base64"), mediaType: "image/jpeg", filename };
    }
    if (mime.startsWith("text/")) {
      return { mode: "text", text: buffer.toString("utf8", 0, Math.min(buffer.length, 80000)), filename };
    }
  }

  // ── Last resort: try reading as UTF-8 text ─────────────────────────
  logger.info({ filename }, "[converter] last resort: raw text read");
  const rawText = buffer.toString("utf8", 0, Math.min(buffer.length, 50000));
  const printableRatio = rawText.split("").filter(c => c.charCodeAt(0) > 31 && c.charCodeAt(0) < 127).length / Math.max(rawText.length, 1);
  if (printableRatio > 0.7) {
    return { mode: "text", text: rawText, filename };
  }

  // Really last resort: tell Claude what we know
  return {
    mode: "text",
    text: `[Binary file: ${filename}, ${buffer.length} bytes, extension: .${e}. Could not auto-convert. Please process based on filename and context clues only.]`,
    filename,
  };
}
