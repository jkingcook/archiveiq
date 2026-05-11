import { Router } from "express";
import multer from "multer";
import Anthropic from "@anthropic-ai/sdk";
import { v4 as uuidv4 } from "uuid";
import { join } from "path";
import { BUS, updatePatternLibrary, isMachenProject } from "../lib/intelligence-bus.js";
import { buildMachenItemDocx, buildGenericItemDocx } from "../lib/docx-builder.js";
import type { MachensItem, GenericItem, ArchiveItem } from "../lib/intelligence-bus.js";

const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }

function imageMediaType(filename: string): "image/jpeg" | "image/png" | "image/webp" | "image/gif" {
  const ext = filename.toLowerCase().split(".").pop();
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "png") return "image/png";
  if (ext === "webp") return "image/webp";
  return "image/jpeg";
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
  const base64 = fileBuffer.toString("base64");
  const mediaType = imageMediaType(filename);

  const systemPrompt = `You are ArchiveIQ, an expert archival document intelligence system. You are processing a document for PROJECT: ${project.name}. You apply the project's Processing Protocol exactly. You return ONLY valid JSON with no markdown, no code fences, no preamble whatsoever.`;

  const userText = `Document image attached.
Project: ${project.name}
Project type: ${project.type}
Source year / date: ${sourceYear || "unknown"}
Source label: ${sourceLabel || filename}
Additional notes: ${notes || "none"}

Processing Protocol (apply exactly):
${protocol}

Output Schema (fill every field):
${itemSchema}

Instructions:
1. Identify the document type from the image
2. Apply every step of the Processing Protocol
${isMachen ? `3. Apply all three passes; build register_rows at day-and-event granularity; Voice rows = direct quotation; apply donee-locator priority; append donee-anchor citations where substantively important` : `3. Extract all people, places, dates, organizations into their arrays`}
4. Set noteworthy_flag = true for items of unusual historical significance
5. Assign confidence_score based on image legibility (high/medium/low)
6. Write bibliography_entry in Chicago citation style
7. Note all OCR corrections in ocr_correction_notes
8. Return ONLY the JSON. No other text.`;

  let parsed: Record<string, unknown> | null = null;
  let rawText = "";

  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const response = await client.messages.create({
        model: "claude-sonnet-4-5",
        max_tokens: 4000,
        system: systemPrompt,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
            { type: "text", text: userText }
          ],
        }],
      });
      rawText = response.content[0].type === "text" ? response.content[0].text : "";
      const cleaned = rawText.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
      parsed = JSON.parse(cleaned);
      break;
    } catch (_e) {
      if (attempt === 0) await sleep(2000);
    }
  }

  const itemId = uuidv4();
  const now = new Date().toISOString();

  if (!parsed) {
    const fallback = {
      item_id: itemId,
      project_id: projectId,
      raw_text: rawText,
      needs_review: true,
      status: "needs_review",
      filename,
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
  parsed.processed_at = now;
  parsed.status = "done";

  const item = parsed as unknown as ArchiveItem;
  BUS.itemStore.push(item);
  updatePatternLibrary(item);

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
  } catch (_e) {
    // docx generation failed — item still saved
  }

  return item;
}

router.post("/", upload.array("files", 50), async (req, res) => {
  const { projectId, sourceYear, sourceLabel, notes } = req.body;
  if (!projectId || !BUS.projects[projectId]) {
    return res.status(400).json({ error: "Valid projectId required" });
  }

  const files = (req.files as Express.Multer.File[]) ?? [];
  if (files.length === 0) return res.status(400).json({ error: "No files provided" });

  if (!process.env["ANTHROPIC_API_KEY"]) {
    return res.status(400).json({ error: "ANTHROPIC_API_KEY not set" });
  }

  const results: ArchiveItem[] = [];
  const errors: string[] = [];

  for (const file of files) {
    try {
      const item = await processOneItem(file.buffer, file.originalname, projectId, sourceYear ?? "", sourceLabel ?? "", notes ?? "");
      results.push(item);
    } catch (e) {
      errors.push(`${file.originalname}: ${String(e)}`);
      BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "process_error", project_id: projectId, details: String(e) });
    }
    await sleep(1000);
  }

  res.json({ success: true, processed: results.length, errors, items: results });
});

router.get("/", (_req, res) => {
  res.json({ items: BUS.itemStore });
});

router.get("/:projectId", (req, res) => {
  const items = BUS.itemStore.filter(i => i.project_id === req.params.projectId);
  res.json({ items });
});

router.post("/search", (req, res) => {
  const { query, projectId, noteworthy_only, confidence } = req.body;
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
