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

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
});

const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }

async function callClaude(
  systemPrompt: string,
  userText: string,
  converted: Awaited<ReturnType<typeof convertFile>>,
  maxTokens: number
): Promise<string> {
  type ContentBlock =
    | { type: "text"; text: string }
    | { type: "image"; source: { type: "base64"; media_type: string; data: string } }
    | { type: "document"; source: { type: "base64"; media_type: "application/pdf"; data: string } };

  let content: ContentBlock[];

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
    // text mode — no image attachment
    const combinedText = `Document text extracted from: ${converted.filename}\n\n${
      converted.mode === "text" ? converted.text : ""
    }\n\n---\n\n${userText}`;
    content = [{ type: "text", text: combinedText }];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const params: any = {
    model: "claude-sonnet-4-5",
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [{ role: "user", content }],
  };

  // Add PDF beta header if using document type
  if (converted.mode === "pdf-doc") {
    params.betas = ["pdfs-2024-09-25"];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const response: any = converted.mode === "pdf-doc"
    ? await (client.beta as any).messages.create(params)
    : await client.messages.create(params);

  const rawText = response.content?.[0]?.type === "text" ? response.content[0].text : "";
  return rawText;
}

async function processOneItem(
  fileBuffer: Buffer,
  filename: string,
  projectId: string,
  sourceYear: string,
  sourceLabel: string,
  notes: string
): Promise<ArchiveItem> {
  const project = BUS.projects[projectId];
  if (!project) throw new Error("Project not found");

  const protocols = BUS.protocolRegistry[projectId] ?? [];
  const schemas = BUS.schemaRegistry[projectId] ?? [];
  const protocol = protocols[protocols.length - 1]?.content ?? "Extract all relevant information from this document.";
  const itemSchema = schemas[schemas.length - 1]?.item_schema ?? "{}";

  const isMachen = isMachenProject(projectId);

  // Convert file to processable form
  const converted = await convertFile(fileBuffer, filename);
  logger.info({ filename, mode: converted.mode }, "[process] file converted");

  // Handle multi-text (zip contents, multi-page text)
  if (converted.mode === "multi-text") {
    // Process first text chunk, merge rows later if needed
    const firstConverted = { mode: "text" as const, text: converted.texts.join("\n\n---\n\n"), filename };
    return processConverted(firstConverted, filename, projectId, project, protocol, itemSchema, isMachen, sourceYear, sourceLabel, notes);
  }

  return processConverted(converted, filename, projectId, project, protocol, itemSchema, isMachen, sourceYear, sourceLabel, notes);
}

async function processConverted(
  converted: Awaited<ReturnType<typeof convertFile>>,
  filename: string,
  projectId: string,
  project: { name: string; type: string },
  protocol: string,
  itemSchema: string,
  isMachen: boolean,
  sourceYear: string,
  sourceLabel: string,
  notes: string
): Promise<ArchiveItem> {
  const systemPrompt = `You are ArchiveIQ, an expert archival document intelligence system. You are processing a document for PROJECT: ${project.name}. You apply the project's Processing Protocol exactly. You return ONLY valid JSON with no markdown, no code fences, no preamble whatsoever.`;

  const modeLabel = converted.mode === "image"
    ? `image (${converted.mediaType})`
    : converted.mode === "pdf-doc"
    ? "PDF document (scanned/image-based)"
    : `extracted text from ${filename}`;

  const userText = `Document: ${filename}
Format: ${modeLabel}
Project: ${project.name}
Project type: ${project.type}
Source year / date: ${sourceYear || "unknown"}
Source label: ${sourceLabel || filename}
Additional notes: ${notes || "none"}

Processing Protocol (apply exactly):
${protocol.substring(0, 5000)}

Output Schema (fill every field):
${itemSchema.substring(0, 3000)}

Instructions:
1. Identify the document type from the content
2. Apply every step of the Processing Protocol
${isMachen
    ? `3. Apply all three passes; build register_rows at day-and-event granularity; Voice rows = direct quotation; apply donee-locator priority; append donee-anchor citations where substantively important`
    : `3. Extract all people, places, dates, organizations into their arrays`
}
4. Set noteworthy_flag = true for items of unusual historical significance
5. Assign confidence_score based on legibility / extraction quality (high/medium/low)
6. Write bibliography_entry in Chicago citation style
7. Note all OCR corrections in ocr_correction_notes
8. Return ONLY the JSON object. No other text whatsoever.`;

  let parsed: Record<string, unknown> | null = null;
  let rawText = "";

  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      rawText = await callClaude(systemPrompt, userText, converted, 4000);
      const cleaned = rawText.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
      // Find the first { to handle any leading whitespace/text
      const jsonStart = cleaned.indexOf("{");
      const jsonEnd = cleaned.lastIndexOf("}");
      if (jsonStart !== -1 && jsonEnd !== -1) {
        parsed = JSON.parse(cleaned.substring(jsonStart, jsonEnd + 1));
      } else {
        parsed = JSON.parse(cleaned);
      }
      break;
    } catch (e) {
      logger.warn({ filename, attempt, err: String(e) }, "[process] Claude call or parse failed");
      if (attempt === 0) await sleep(2000);
    }
  }

  const itemId = uuidv4();
  const now = new Date().toISOString();

  if (!parsed) {
    logger.error({ filename }, "[process] both attempts failed → creating needs_review item");
    const fallback = {
      item_id: itemId,
      project_id: projectId,
      raw_text: rawText.substring(0, 500),
      needs_review: true,
      status: "needs_review",
      filename,
      file_mode: converted.mode,
      processed_at: now,
      people_extracted: [],
      places_extracted: [],
      dates_extracted: [],
      register_rows: [],
      noteworthy_flag: false,
      confidence_score: "low",
    } as unknown as ArchiveItem;
    BUS.itemStore.push(fallback);
    updatePatternLibrary(fallback);
    return fallback;
  }

  parsed.item_id = itemId;
  parsed.project_id = projectId;
  parsed.filename = filename;
  parsed.file_mode = converted.mode;
  parsed.processed_at = now;
  parsed.status = "done";

  const item = parsed as unknown as ArchiveItem;
  BUS.itemStore.push(item);
  updatePatternLibrary(item);

  // Generate .docx
  try {
    const safeName = filename.replace(/[^a-zA-Z0-9._-]/g, "_");
    let outPath: string;
    if (isMachen) {
      const year = sourceYear || (item as MachensItem).source_year || "unknown";
      outPath = `/tmp/output/items/Biographical_Facts_Register_Diary_${year}_${itemId.substring(0, 8)}.docx`;
      await buildMachenItemDocx(item as MachensItem, outPath);
    } else {
      outPath = `/tmp/output/items/${project.name.replace(/\s+/g, "_")}_${safeName}_${itemId.substring(0, 8)}.docx`;
      await buildGenericItemDocx(item, outPath);
    }
    (item as { docx_path?: string }).docx_path = outPath;
  } catch (e) {
    logger.warn({ filename, err: String(e) }, "[process] docx generation failed (non-fatal)");
  }

  return item;
}

// ── Routes ────────────────────────────────────────────────────────────

router.post("/", upload.array("files", 50), async (req: Request, res: Response) => {
  const { projectId, sourceYear, sourceLabel, notes } = req.body as Record<string, string>;
  if (!projectId || !BUS.projects[projectId]) {
    return res.status(400).json({ error: "Valid projectId required" });
  }

  const files = (req.files as Express.Multer.File[]) ?? [];
  if (files.length === 0) {
    return res.status(400).json({ error: "No files provided" });
  }

  if (!process.env["ANTHROPIC_API_KEY"]) {
    return res.status(400).json({ error: "ANTHROPIC_API_KEY not set — add it to Replit Secrets" });
  }

  const results: ArchiveItem[] = [];
  const errors: string[] = [];

  for (const file of files) {
    try {
      const item = await processOneItem(
        file.buffer,
        file.originalname,
        projectId,
        sourceYear ?? "",
        sourceLabel ?? "",
        notes ?? ""
      );
      results.push(item);
    } catch (e) {
      const msg = `${file.originalname}: ${String(e)}`;
      errors.push(msg);
      logger.error({ filename: file.originalname, err: String(e) }, "[process] item processing error");
      BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "process_error", project_id: projectId, details: msg });
    }
    await sleep(1000);
  }

  res.json({ success: true, processed: results.length, errors, items: results });
});

// Multer error handler — returns JSON instead of HTML
router.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  logger.error({ err: err.message }, "[process] middleware error");
  res.status(400).json({ error: err.message ?? "File upload error" });
});

router.get("/", (_req: Request, res: Response) => {
  res.json({ items: BUS.itemStore });
});

router.get("/:projectId", (req: Request, res: Response) => {
  const items = BUS.itemStore.filter(i => i.project_id === req.params["projectId"]);
  res.json({ items });
});

router.post("/search", (req: Request, res: Response) => {
  const { query, projectId, noteworthy_only, confidence } = req.body as Record<string, string>;
  let items = projectId ? BUS.itemStore.filter(i => i.project_id === projectId) : BUS.itemStore;
  if (noteworthy_only) items = items.filter(i => (i as { noteworthy_flag?: boolean }).noteworthy_flag);
  if (confidence) items = items.filter(i => (i as { confidence_score?: string }).confidence_score === confidence);
  if (query) {
    const q = String(query).toLowerCase();
    items = items.filter(i => JSON.stringify(i).toLowerCase().includes(q));
  }
  res.json({ items, count: items.length });
});

export default router;
