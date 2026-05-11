import { Router, type Request, type Response, type NextFunction } from "express";
import multer from "multer";
import Anthropic from "@anthropic-ai/sdk";
import { v4 as uuidv4 } from "uuid";
import { BUS, updatePatternLibrary, isMachenProject } from "../lib/intelligence-bus.js";
import { buildMachenItemDocx, buildGenericItemDocx } from "../lib/docx-builder.js";
import { convertFile } from "../lib/file-converter.js";
import { logger } from "../lib/logger.js";
import type { MachensItem, ArchiveItem } from "../lib/intelligence-bus.js";

const router = Router();

// No size or count limits
const upload = multer({ storage: multer.memoryStorage() });

const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

const PAGES_PER_CHUNK = 5;
const CHARS_PER_CHUNK = 20000;

type ProgressFn = (chunk: number, totalChunks: number, page: number, totalPages: number) => void;

type ProcessableFile =
  | { mode: "image"; base64: string; mediaType: string; filename: string }
  | { mode: "pdf-doc"; base64: string; filename: string }
  | { mode: "text"; text: string; filename: string };

function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }

function parseJSON(raw: string): Record<string, unknown> | null {
  try {
    const cleaned = raw.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
    const s = cleaned.indexOf("{");
    const e = cleaned.lastIndexOf("}");
    if (s !== -1 && e !== -1) return JSON.parse(cleaned.substring(s, e + 1));
    return JSON.parse(cleaned);
  } catch { return null; }
}

function getTextChunks(text: string): string[] {
  const pages = text.split("\f").filter(p => p.trim().length > 10);
  if (pages.length > PAGES_PER_CHUNK) {
    const chunks: string[] = [];
    for (let i = 0; i < pages.length; i += PAGES_PER_CHUNK) {
      const group = pages.slice(i, i + PAGES_PER_CHUNK);
      const range = `Pages ${i + 1}–${Math.min(i + PAGES_PER_CHUNK, pages.length)} of ${pages.length}`;
      chunks.push(`[${range}]\n\n${group.join("\n\n--- Page Break ---\n\n")}`);
    }
    return chunks;
  }
  if (text.length > CHARS_PER_CHUNK) {
    const total = Math.ceil(text.length / CHARS_PER_CHUNK);
    const chunks: string[] = [];
    for (let i = 0; i < text.length; i += CHARS_PER_CHUNK) {
      const n = Math.floor(i / CHARS_PER_CHUNK) + 1;
      chunks.push(`[Section ${n} of ${total}]\n\n${text.slice(i, i + CHARS_PER_CHUNK)}`);
    }
    return chunks;
  }
  return [text];
}

async function callClaude(systemPrompt: string, userText: string, converted: ProcessableFile, maxTokens: number): Promise<string> {
  type Block =
    | { type: "text"; text: string }
    | { type: "image"; source: { type: "base64"; media_type: string; data: string } }
    | { type: "document"; source: { type: "base64"; media_type: "application/pdf"; data: string } };

  let content: Block[];
  if (converted.mode === "image") {
    content = [
      { type: "image", source: { type: "base64", media_type: converted.mediaType, data: converted.base64 } },
      { type: "text", text: userText },
    ];
  } else if (converted.mode === "pdf-doc") {
    content = [
      { type: "document", source: { type: "base64", media_type: "application/pdf", data: converted.base64 } },
      { type: "text", text: userText },
    ];
  } else {
    content = [{ type: "text", text: `Document: ${converted.filename}\n\n${converted.text}\n\n---\n\n${userText}` }];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const params: any = { model: "claude-sonnet-4-5", max_tokens: maxTokens, system: systemPrompt, messages: [{ role: "user", content }] };
  if (converted.mode === "pdf-doc") params.betas = ["pdfs-2024-09-25"];

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const resp: any = converted.mode === "pdf-doc"
    ? await (client.beta as any).messages.create(params)
    : await client.messages.create(params);

  return resp.content?.[0]?.type === "text" ? resp.content[0].text : "";
}

function createNeedsReviewItem(projectId: string, filename: string, rawText: string): ArchiveItem {
  const item = {
    item_id: uuidv4(), project_id: projectId, raw_text: rawText, needs_review: true,
    status: "needs_review", filename, processed_at: new Date().toISOString(),
    people_extracted: [], places_extracted: [], dates_extracted: [],
    register_rows: [], noteworthy_flag: false, confidence_score: "low",
  } as unknown as ArchiveItem;
  BUS.itemStore.push(item);
  updatePatternLibrary(item);
  return item;
}

async function finalizeItem(
  parsed: Record<string, unknown>, projectId: string, filename: string,
  isMachen: boolean, project: { name: string; type: string }, sourceYear: string
): Promise<ArchiveItem> {
  const itemId = uuidv4();
  Object.assign(parsed, { item_id: itemId, project_id: projectId, filename, processed_at: new Date().toISOString(), status: "done" });
  const item = parsed as unknown as ArchiveItem;
  BUS.itemStore.push(item);
  updatePatternLibrary(item);

  try {
    const safe = filename.replace(/[^a-zA-Z0-9._-]/g, "_");
    let outPath: string;
    if (isMachen) {
      const year = sourceYear || (item as MachensItem).source_year || "unknown";
      outPath = `/tmp/output/items/Biographical_Facts_Register_Diary_${year}_${itemId.slice(0, 8)}.docx`;
      await buildMachenItemDocx(item as MachensItem, outPath);
    } else {
      outPath = `/tmp/output/items/${project.name.replace(/\s+/g, "_")}_${safe}_${itemId.slice(0, 8)}.docx`;
      await buildGenericItemDocx(item, outPath);
    }
    (item as { docx_path?: string }).docx_path = outPath;
  } catch (e) {
    logger.warn({ filename, err: String(e) }, "[process] docx failed (non-fatal)");
  }
  return item;
}

async function processConverted(
  converted: ProcessableFile,
  filename: string,
  projectId: string,
  project: { name: string; type: string },
  protocol: string,
  itemSchema: string,
  isMachen: boolean,
  sourceYear: string,
  sourceLabel: string,
  notes: string,
  sendProgress: ProgressFn
): Promise<ArchiveItem> {
  const system = `You are ArchiveIQ, expert archival intelligence for PROJECT: ${project.name}. Apply the Processing Protocol exactly. Return ONLY valid JSON — no markdown, no fences, no preamble.`;

  const modeLabel = converted.mode === "image"
    ? `image (${converted.mediaType})`
    : converted.mode === "pdf-doc" ? "PDF (scanned/image-based)" : `extracted text from ${filename}`;

  const baseUser = `Document: ${filename}
Format: ${modeLabel}
Project: ${project.name} (${project.type})
Source year: ${sourceYear || "unknown"}
Source label: ${sourceLabel || filename}
Notes: ${notes || "none"}

Processing Protocol:
${protocol.substring(0, 5000)}

Output Schema:
${itemSchema.substring(0, 3000)}

Instructions:
1. Apply every step of the Processing Protocol
${isMachen ? "2. Apply all three passes; build register_rows at day-and-event granularity; Voice rows = direct quotation; apply donee-locator priority" : "2. Extract all people, places, dates, organizations"}
3. noteworthy_flag=true for unusual historical significance
4. confidence_score: high/medium/low based on legibility
5. bibliography_entry in Chicago style
6. Record OCR corrections in ocr_correction_notes
7. Return ONLY the JSON object.`;

  // Determine chunks (only for text mode)
  const textChunks = converted.mode === "text" ? getTextChunks(converted.text) : [];
  const isMulti = textChunks.length > 1;

  if (!isMulti) {
    sendProgress(1, 1, 1, 1);
    let raw = "";
    let parsed: Record<string, unknown> | null = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        raw = await callClaude(system, baseUser, converted, 4000);
        parsed = parseJSON(raw);
        if (parsed) break;
      } catch (e) {
        logger.warn({ filename, attempt, err: String(e) }, "[process] call failed");
        if (attempt === 0) await sleep(2000);
      }
    }
    if (!parsed) return createNeedsReviewItem(projectId, filename, raw.substring(0, 500));
    return finalizeItem(parsed, projectId, filename, isMachen, project, sourceYear);
  }

  // Multi-chunk
  logger.info({ filename, chunks: textChunks.length }, "[process] chunked processing");
  const allRows: object[] = [];
  const people = new Set<string>();
  const places = new Set<string>();
  const dates = new Set<string>();
  let base: Record<string, unknown> | null = null;

  for (let ci = 0; ci < textChunks.length; ci++) {
    const estPage = ci * PAGES_PER_CHUNK + 1;
    const estTotal = textChunks.length * PAGES_PER_CHUNK;
    sendProgress(ci + 1, textChunks.length, estPage, estTotal);

    const chunkFile: ProcessableFile = { mode: "text", text: textChunks[ci], filename };
    const suffix = ci === 0
      ? `\n\n[Chunk 1/${textChunks.length} — provide complete schema output]`
      : `\n\n[Chunk ${ci + 1}/${textChunks.length} — return JSON with ONLY: register_rows, people_extracted, places_extracted, dates_extracted for this section]`;

    let parsed: Record<string, unknown> | null = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const raw = await callClaude(system, baseUser + suffix, chunkFile, 4000);
        parsed = parseJSON(raw);
        if (parsed) break;
      } catch (e) {
        logger.warn({ filename, chunk: ci + 1, attempt, err: String(e) }, "[process] chunk failed");
        if (attempt === 0) await sleep(2000);
      }
    }

    if (parsed) {
      if (ci === 0 || !base) base = { ...parsed };
      ((parsed["register_rows"] as object[]) || []).forEach(r => allRows.push(r));
      ((parsed["people_extracted"] as string[]) || []).forEach(p => people.add(p));
      ((parsed["places_extracted"] as string[]) || []).forEach(p => places.add(p));
      ((parsed["dates_extracted"] as string[]) || []).forEach(d => dates.add(d));
    }

    if (ci < textChunks.length - 1) await sleep(1000);
  }

  if (!base) return createNeedsReviewItem(projectId, filename, "");

  Object.assign(base, {
    register_rows: allRows,
    people_extracted: [...people],
    places_extracted: [...places],
    dates_extracted: [...dates],
    chunk_count: textChunks.length,
  });

  return finalizeItem(base, projectId, filename, isMachen, project, sourceYear);
}

async function processOneItem(
  buf: Buffer, filename: string, projectId: string,
  sourceYear: string, sourceLabel: string, notes: string,
  sendProgress: ProgressFn
): Promise<ArchiveItem> {
  const project = BUS.projects[projectId];
  if (!project) throw new Error("Project not found");

  const protocols = BUS.protocolRegistry[projectId] ?? [];
  const schemas = BUS.schemaRegistry[projectId] ?? [];
  const protocol = protocols[protocols.length - 1]?.content ?? "Extract all relevant information from this document.";
  const itemSchema = schemas[schemas.length - 1]?.item_schema ?? "{}";
  const isMachen = isMachenProject(projectId);

  const converted = await convertFile(buf, filename);
  logger.info({ filename, mode: converted.mode }, "[process] converted");

  // Normalize to ProcessableFile
  let pf: ProcessableFile;
  if (converted.mode === "image") {
    pf = { mode: "image", base64: converted.base64, mediaType: converted.mediaType, filename };
  } else if (converted.mode === "pdf-doc") {
    pf = { mode: "pdf-doc", base64: converted.base64, filename };
  } else if (converted.mode === "text") {
    pf = { mode: "text", text: converted.text, filename };
  } else if (converted.mode === "multi-text") {
    pf = { mode: "text", text: converted.texts.join("\n\n---\n\n"), filename };
  } else {
    pf = { mode: "text", text: `[Unsupported format: ${filename}]`, filename };
  }

  return processConverted(pf, filename, projectId, project, protocol, itemSchema, isMachen, sourceYear, sourceLabel, notes, sendProgress);
}

// ── SSE Route ────────────────────────────────────────────────────────────

router.post("/", upload.array("files"), async (req: Request, res: Response) => {
  const { projectId, sourceYear, sourceLabel, notes } = req.body as Record<string, string>;

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

  const files = (req.files as Express.Multer.File[]) ?? [];
  if (files.length === 0) {
    send({ type: "error", error: "No files provided" });
    return res.end();
  }

  if (!process.env["ANTHROPIC_API_KEY"]) {
    send({ type: "error", error: "ANTHROPIC_API_KEY not set" });
    return res.end();
  }

  const results: ArchiveItem[] = [];
  const errors: string[] = [];

  for (const file of files) {
    send({ type: "start", filename: file.originalname, fileIndex: files.indexOf(file), totalFiles: files.length });
    try {
      const item = await processOneItem(
        file.buffer, file.originalname, projectId,
        sourceYear ?? "", sourceLabel ?? "", notes ?? "",
        (chunk, totalChunks, page, totalPages) =>
          send({ type: "progress", filename: file.originalname, chunk, totalChunks, page, totalPages })
      );
      results.push(item);
      send({ type: "file_done", filename: file.originalname, itemId: item.item_id, status: item.status, needs_review: (item as { needs_review?: boolean }).needs_review });
    } catch (e) {
      const msg = `${file.originalname}: ${String(e)}`;
      errors.push(msg);
      logger.error({ filename: file.originalname, err: String(e) }, "[process] error");
      BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "process_error", project_id: projectId, details: msg });
      send({ type: "file_error", filename: file.originalname, error: String(e) });
    }
    await sleep(1000);
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

router.get("/", (_req: Request, res: Response) => {
  res.json({ items: BUS.itemStore });
});

router.get("/:projectId", (req: Request, res: Response) => {
  res.json({ items: BUS.itemStore.filter(i => i.project_id === req.params["projectId"]) });
});

router.post("/search", (req: Request, res: Response) => {
  const { query, projectId, noteworthy_only, confidence } = req.body as Record<string, string>;
  let items = projectId ? BUS.itemStore.filter(i => i.project_id === projectId) : BUS.itemStore;
  if (noteworthy_only) items = items.filter(i => (i as { noteworthy_flag?: boolean }).noteworthy_flag);
  if (confidence) items = items.filter(i => (i as { confidence_score?: string }).confidence_score === confidence);
  if (query) { const q = query.toLowerCase(); items = items.filter(i => JSON.stringify(i).toLowerCase().includes(q)); }
  res.json({ items, count: items.length });
});

export default router;
