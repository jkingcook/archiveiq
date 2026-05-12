import { Router, type Request, type Response, type NextFunction } from "express";
import multer from "multer";
import Anthropic from "@anthropic-ai/sdk";
import { createRequire } from "module";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { v4 as uuidv4 } from "uuid";
import { BUS } from "../lib/intelligence-bus.js";
import { logger } from "../lib/logger.js";

const _require = createRequire(import.meta.url);
const router = Router();
const upload = multer({ storage: multer.memoryStorage() });
const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }

// ── Text extraction ───────────────────────────────────────────────────────────

async function extractText(buffer: Buffer, filename: string): Promise<string> {
  const ext = (filename.split(".").pop() ?? "").toLowerCase();

  if (ext === "pdf") {
    // pdf-parse v2: requires file:// URL — write buffer to tmp file, parse, clean up
    const tmpPath = join(tmpdir(), `aiq_${Date.now()}_${uuidv4().slice(0, 8)}.pdf`);
    try {
      writeFileSync(tmpPath, buffer);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { PDFParse, VerbosityLevel } = _require("pdf-parse") as any;
      const parser = new PDFParse({ url: `file://${tmpPath}`, verbosity: VerbosityLevel.ERRORS });
      const result = await parser.getText() as { text?: string; numpages?: number };
      const text = result.text?.trim() ?? "";
      logger.info({ filename, textLen: text.length, textPreview: text.substring(0, 200) }, "[extract] pdf-parse ok");
      return text;
    } catch (e) {
      logger.error({ filename, err: String(e) }, "[extract] pdf-parse failed");
      return "";
    } finally {
      try { unlinkSync(tmpPath); } catch { /* ignore */ }
    }
  }

  if (["txt", "md", "csv", "tsv"].includes(ext)) {
    return buffer.toString("utf-8").trim();
  }

  if (["docx"].includes(ext)) {
    try {
      const mammoth = _require("mammoth");
      const result = await mammoth.extractRawText({ buffer });
      return (result.value as string).trim();
    } catch (e) {
      logger.warn({ filename, err: String(e) }, "[extract] mammoth failed");
      return "";
    }
  }

  return "";
}

// ── Claude call ───────────────────────────────────────────────────────────────

async function callClaude(buffer: Buffer, filename: string): Promise<Record<string, unknown>> {
  const ext = (filename.split(".").pop() ?? "").toLowerCase();
  const isImage = ["jpg", "jpeg", "png", "gif", "webp"].includes(ext);

  const SYSTEM = "You process archival documents. Return ONLY valid JSON. No markdown. No explanation.";

  const SCHEMA = `{
  "item_type": "string",
  "full_transcription": "string",
  "people_extracted": ["string"],
  "places_extracted": ["string"],
  "dates_extracted": ["string"],
  "register_rows": [{"date": "string", "person": "string", "category": "string", "fact": "string", "source": "string"}],
  "confidence_score": "high | medium | low"
}`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let content: any[];

  if (isImage) {
    const mediaType = (ext === "jpg" || ext === "jpeg") ? "image/jpeg" : `image/${ext}`;
    const base64 = buffer.toString("base64");
    content = [
      { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
      { type: "text", text: `Process this archival document image and return a JSON object with these exact fields:\n${SCHEMA}` },
    ];
    logger.info({ filename, mode: "image", mediaType }, "[claude] sending image");
  } else {
    const text = await extractText(buffer, filename);
    logger.info({ filename, textLen: text.length, textPreview: text.substring(0, 500) }, "[extract] text ready");

    if (!text) {
      throw new Error(`Could not extract text from ${filename}`);
    }

    const userMsg = `Process this document text and return a JSON object with these exact fields:\n${SCHEMA}\n\nDocument text:\n${text}`;
    content = [{ type: "text", text: userMsg }];
    logger.info({ filename, mode: "text", promptLen: userMsg.length, max_tokens: 8096 }, "[claude] sending request");
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const resp: any = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 8096,
    system: SYSTEM,
    messages: [{ role: "user", content }],
  });

  const raw: string = resp.content?.[0]?.type === "text" ? (resp.content[0].text as string) : "";
  const stopReason = (resp.stop_reason as string) ?? "end_turn";
  const outputTokens = (resp.usage?.output_tokens as number) ?? 0;

  logger.info({
    filename,
    stopReason,
    outputTokens,
    responseLen: raw.length,
    rawPreview: raw.substring(0, 500),
  }, "[claude] response received");

  if (stopReason === "max_tokens") {
    logger.error({ filename, outputTokens }, "[claude] TRUNCATED by max_tokens");
  }

  // ── Parse ─────────────────────────────────────────────────────────────────
  const cleaned = raw
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```\s*$/, "")
    .trim();

  const s = cleaned.indexOf("{");
  const e = cleaned.lastIndexOf("}");

  if (s === -1) {
    logger.error({ filename, rawFull: raw }, "[parse] FAILED — no JSON object in response (raw logged above)");
    throw new Error("Claude did not return a JSON object");
  }

  const candidate = e > s ? cleaned.substring(s, e + 1) : cleaned.substring(s);

  try {
    const parsed = JSON.parse(candidate) as Record<string, unknown>;
    const rowCount = (parsed["register_rows"] as unknown[])?.length ?? 0;
    logger.info({ filename, rowCount, keys: Object.keys(parsed) }, "[parse] SUCCESS");
    return parsed;
  } catch (firstErr) {
    // Truncation rescue: close unclosed brackets/braces
    logger.warn({ filename, err: String(firstErr), rawTail: raw.slice(-300) }, "[parse] first attempt failed — trying truncation rescue");
    let rescued = candidate;
    const openBraces = (rescued.match(/\{/g) ?? []).length - (rescued.match(/\}/g) ?? []).length;
    const openArrays = (rescued.match(/\[/g) ?? []).length - (rescued.match(/\]/g) ?? []).length;
    const lastChar = rescued.trimEnd().slice(-1);
    if (lastChar !== '"' && lastChar !== "}" && lastChar !== "]") rescued += '"';
    for (let i = 0; i < openArrays; i++) rescued += "]";
    for (let i = 0; i < openBraces; i++) rescued += "}";

    try {
      const parsed = JSON.parse(rescued) as Record<string, unknown>;
      const rowCount = (parsed["register_rows"] as unknown[])?.length ?? 0;
      logger.info({ filename, rowCount }, "[parse] truncation rescue SUCCESS");
      return parsed;
    } catch (rescueErr) {
      logger.error({ filename, firstErr: String(firstErr), rescueErr: String(rescueErr), rawFull: raw }, "[parse] FAILED ALL ATTEMPTS — full raw response logged above");
      throw new Error(`JSON parse failed: ${String(firstErr)}`);
    }
  }
}

// ── SSE route ─────────────────────────────────────────────────────────────────

router.post("/", upload.array("files"), async (req: Request, res: Response) => {
  const { projectId } = req.body as Record<string, string>;

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("X-Accel-Buffering", "no");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();

  const send = (data: object) => {
    if (!res.writableEnded) res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  if (!projectId || !BUS.projects[projectId]) {
    send({ type: "error", error: "Valid projectId required" });
    return res.end();
  }

  if (!process.env["ANTHROPIC_API_KEY"]) {
    send({ type: "error", error: "ANTHROPIC_API_KEY not set" });
    return res.end();
  }

  const files = (req.files as Express.Multer.File[]) ?? [];
  if (files.length === 0) {
    send({ type: "error", error: "No files provided" });
    return res.end();
  }

  const results: object[] = [];
  const errors: string[] = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    send({ type: "start", filename: file.originalname, fileIndex: i, totalFiles: files.length });
    send({ type: "progress", filename: file.originalname, chunk: 1, totalChunks: 1, page: 1, totalPages: 1 });

    try {
      const parsed = await callClaude(file.buffer, file.originalname);

      const itemId = uuidv4();
      const item = {
        ...parsed,
        item_id: itemId,
        project_id: projectId,
        filename: file.originalname,
        processed_at: new Date().toISOString(),
        status: "done",
        needs_review: false,
      };

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      BUS.itemStore.push(item as any);
      results.push(item);

      const rowCount = (parsed["register_rows"] as unknown[])?.length ?? 0;
      logger.info({ filename: file.originalname, itemId, rowCount }, "[process] item stored");

      send({ type: "file_done", filename: file.originalname, itemId, status: "done", needs_review: false });

    } catch (e) {
      const msg = `${file.originalname}: ${String(e)}`;
      errors.push(msg);
      logger.error({ filename: file.originalname, err: String(e) }, "[process] file failed");

      const itemId = uuidv4();
      const fallback = {
        item_id: itemId,
        project_id: projectId,
        filename: file.originalname,
        processed_at: new Date().toISOString(),
        status: "needs_review",
        needs_review: true,
        register_rows: [],
        people_extracted: [],
        places_extracted: [],
        dates_extracted: [],
        item_type: "unknown",
        confidence_score: "low",
      };

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      BUS.itemStore.push(fallback as any);
      results.push(fallback);

      send({ type: "file_done", filename: file.originalname, itemId, status: "needs_review", needs_review: true });
    }

    if (i < files.length - 1) await sleep(1000);
  }

  send({ type: "done", processed: results.length, errors, items: results });
  res.end();
});

router.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  logger.error({ err: err.message }, "[process] middleware error");
  if (!res.writableEnded) {
    res.setHeader("Content-Type", "application/json");
    res.status(400).json({ error: err.message ?? "Upload error" });
  }
});

export default router;
