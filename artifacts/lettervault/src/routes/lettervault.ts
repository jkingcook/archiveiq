import { Router, type IRouter, type Request, type Response } from "express";
import multer from "multer";
import { v4 as uuidv4 } from "uuid";
import path from "node:path";
import fs from "node:fs";
import fsp from "node:fs/promises";
import Anthropic from "@anthropic-ai/sdk";
import { ZipArchive } from "archiver";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  Header,
  Footer,
} from "docx";

const router: IRouter = Router();

const ITEMS_DIR = "/tmp/output/items";
const ANALYSIS_DIR = "/tmp/output/analysis";

fs.mkdirSync(ITEMS_DIR, { recursive: true });
fs.mkdirSync(ANALYSIS_DIR, { recursive: true });

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
});

const DEFAULT_ITEM_SCHEMA = JSON.stringify({
  item_id: "string (UUID assigned automatically)",
  item_type: "letter | news_clipping | photograph | handwritten_note | typed_document | postcard | other",
  date_on_item: "string or null",
  date_circa: "boolean",
  sender_name: "string or null",
  recipient_name: "string or null",
  location_origin: "string or null",
  location_mentioned: ["array of place names"],
  people_mentioned: ["array of full names"],
  organizations_mentioned: ["array"],
  subject_summary: "string (2-3 sentences)",
  full_transcription: "string (verbatim text from image)",
  keywords: ["array of topic tags"],
  era_tag: "string (e.g. 1920s, Civil War era, WWII, etc.)",
  language: "string",
  condition_notes: "string (image quality, legibility)",
  noteworthy_flag: "boolean — true if item is historically significant",
  noteworthy_reason: "string or null",
  bibliography_entry: "string — formatted citation for this item",
  protocol_notes: "string — how protocol was applied",
  confidence_score: "high | medium | low",
}, null, 2);

const DEFAULT_ANALYSIS_SCHEMA = JSON.stringify({
  group_label: "string (e.g. '1943' or 'Letters from John')",
  group_criteria: "string (what defined this group)",
  item_count: "number",
  date_range: "string",
  narrative_diary: "string — flowing narrative prose in chronological order",
  key_events: ["array of significant events with dates"],
  people_directory: [{ name: "string", role: "string", first_appearance: "item_id", mention_count: "number" }],
  places_directory: [{ place: "string", context: "string", mention_count: "number" }],
  organizations_directory: ["array"],
  themes: ["array of dominant themes across this group"],
  noteworthy_items: ["array of item_ids flagged as significant"],
  bibliography: ["array of bibliography_entry strings from all items"],
  archivist_summary: "string — 1 paragraph synthesis for researchers",
  gaps_and_unknowns: "string — what is missing or unclear",
}, null, 2);

interface SessionData {
  protocol?: string;
  itemSchema?: string;
  analysisSchema?: string;
}

interface ProcessedItem {
  item_id: string;
  filename: string;
  item_type?: string;
  date_on_item?: string | null;
  date_circa?: boolean;
  sender_name?: string | null;
  recipient_name?: string | null;
  location_origin?: string | null;
  location_mentioned?: string[];
  people_mentioned?: string[];
  organizations_mentioned?: string[];
  subject_summary?: string;
  full_transcription?: string;
  keywords?: string[];
  era_tag?: string;
  language?: string;
  condition_notes?: string;
  noteworthy_flag?: boolean;
  noteworthy_reason?: string | null;
  bibliography_entry?: string;
  protocol_notes?: string;
  confidence_score?: string;
  needs_review?: boolean;
  raw_response?: string;
  processed_at: string;
  docx_path?: string;
}

const itemStore: ProcessedItem[] = [];
const sessionStore: SessionData = {
  itemSchema: DEFAULT_ITEM_SCHEMA,
  analysisSchema: DEFAULT_ANALYSIS_SCHEMA,
};

function getAnthropicClient(): Anthropic {
  const key = process.env["ANTHROPIC_API_KEY"];
  if (!key) throw new Error("ANTHROPIC_API_KEY not set");
  return new Anthropic({ apiKey: key });
}

function getMediaType(filename: string): "image/jpeg" | "image/png" | "image/gif" | "image/webp" {
  const ext = path.extname(filename).toLowerCase();
  if (ext === ".png") return "image/png";
  if (ext === ".webp") return "image/webp";
  if (ext === ".gif") return "image/gif";
  return "image/jpeg";
}

async function generateItemDocx(item: ProcessedItem, outPath: string): Promise<void> {
  const addHeading = (text: string, level: HeadingLevel) =>
    new Paragraph({ text, heading: level, spacing: { before: 200, after: 100 } });

  const addText = (text: string, bold = false) =>
    new Paragraph({
      children: [new TextRun({ text: text || "", bold, size: 22, font: "Georgia" })],
      spacing: { before: 60, after: 60 },
    });

  const addBullet = (text: string) =>
    new Paragraph({
      text: `• ${text}`,
      spacing: { before: 40, after: 40 },
      indent: { left: 400 },
    });

  const metaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Date", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.date_on_item || "Unknown")] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Type", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.item_type || "Unknown")] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Sender", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.sender_name || "Unknown")] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Recipient", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.recipient_name || "Unknown")] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Origin", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.location_origin || "Unknown")] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Language", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.language || "Unknown")] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Era", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.era_tag || "Unknown")] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Confidence", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.confidence_score || "Unknown")] }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Noteworthy", bold: true })] })] }),
          new TableCell({ children: [new Paragraph(item.noteworthy_flag ? "YES" : "No")] }),
          new TableCell({ children: [new Paragraph("")] }),
          new TableCell({ children: [new Paragraph("")] }),
        ],
      }),
    ],
  });

  const noteworthyParagraphs: Paragraph[] = [];
  if (item.noteworthy_flag && item.noteworthy_reason) {
    noteworthyParagraphs.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `NOTEWORTHY ITEM: ${item.noteworthy_reason}`,
            bold: true,
            color: "C9A84C",
            size: 24,
          }),
        ],
        shading: { type: ShadingType.CLEAR, color: "FFF8E1", fill: "FFF8E1" },
        border: {
          top: { style: BorderStyle.SINGLE, size: 6, color: "C9A84C" },
          bottom: { style: BorderStyle.SINGLE, size: 6, color: "C9A84C" },
          left: { style: BorderStyle.SINGLE, size: 6, color: "C9A84C" },
          right: { style: BorderStyle.SINGLE, size: 6, color: "C9A84C" },
        },
        spacing: { before: 200, after: 200 },
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "LetterVault Pro — Archived Item", bold: true, size: 20 }),
                ],
                alignment: AlignmentType.CENTER,
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: `${item.item_type || "Item"} | ID: ${item.item_id} | Confidence: ${item.confidence_score || "Unknown"}`, size: 18 }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: `LetterVault Pro | Processed: ${item.processed_at}`, size: 18 }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        },
        children: [
          addHeading("LetterVault Pro — Archived Item", HeadingLevel.HEADING_1),
          addText(`File: ${item.filename}`, true),

          ...noteworthyParagraphs,

          addHeading("Section 1: Metadata", HeadingLevel.HEADING_2),
          metaTable,

          addHeading("Section 2: Subject Summary", HeadingLevel.HEADING_2),
          addText(item.subject_summary || "No summary available."),

          addHeading("Section 3: Full Transcription", HeadingLevel.HEADING_2),
          ...(item.full_transcription || "No transcription available.").split("\n").map(line =>
            new Paragraph({
              children: [new TextRun({ text: line, font: "Georgia", size: 22 })],
              spacing: { before: 40, after: 40 },
            })
          ),

          addHeading("Section 4: People & Places", HeadingLevel.HEADING_2),
          addText("People Mentioned:", true),
          ...(item.people_mentioned || []).map(p => addBullet(p)),
          addText("Places Mentioned:", true),
          ...(item.location_mentioned || []).map(p => addBullet(p)),
          addText("Organizations:", true),
          ...(item.organizations_mentioned || []).map(o => addBullet(o)),
          addText("Keywords:", true),
          ...(item.keywords || []).map(k => addBullet(k)),

          addHeading("Section 5: Bibliography", HeadingLevel.HEADING_2),
          addText(item.bibliography_entry || "No bibliography entry."),

          addHeading("Section 6: Protocol & Archivist Notes", HeadingLevel.HEADING_2),
          addText(item.protocol_notes || "No protocol notes."),

          ...(item.needs_review ? [
            addText("*** NEEDS REVIEW: Claude response could not be parsed as valid JSON. See raw_response field. ***", true)
          ] : []),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  await fsp.writeFile(outPath, buffer);
}

async function generateGroupDocx(group: Record<string, unknown>, outPath: string): Promise<void> {
  const label = String(group.group_label || "Group");
  const dateRange = String(group.date_range || "");
  const itemCount = Number(group.item_count || 0);
  const addHeading = (text: string, level: HeadingLevel) =>
    new Paragraph({ text, heading: level, spacing: { before: 200, after: 100 } });
  const addText = (text: string, bold = false) =>
    new Paragraph({ children: [new TextRun({ text: text || "", bold, size: 22, font: "Georgia" })], spacing: { before: 60, after: 60 } });
  const addBullet = (text: string) =>
    new Paragraph({ text: `• ${text}`, indent: { left: 400 }, spacing: { before: 40, after: 40 } });

  const keyEvents = Array.isArray(group.key_events) ? group.key_events as string[] : [];
  const themes = Array.isArray(group.themes) ? group.themes as string[] : [];
  const noteworthy = Array.isArray(group.noteworthy_items) ? group.noteworthy_items as string[] : [];
  const bibliography = Array.isArray(group.bibliography) ? group.bibliography as string[] : [];
  const peopleDir = Array.isArray(group.people_directory) ? group.people_directory as Record<string, string>[] : [];
  const placesDir = Array.isArray(group.places_directory) ? group.places_directory as Record<string, string>[] : [];
  const orgsDir = Array.isArray(group.organizations_directory) ? group.organizations_directory as string[] : [];

  void dateRange;
  void itemCount;

  const doc = new Document({
    sections: [{
      children: [
        addHeading(`Collection Analysis: ${label}`, HeadingLevel.HEADING_1),
        addText(`Date Range: ${String(group.date_range || "")} | Items: ${String(group.item_count || 0)}`, true),

        addHeading("Section 1: Archivist Summary", HeadingLevel.HEADING_2),
        addText(String(group.archivist_summary || "")),

        addHeading("Section 2: Narrative Diary", HeadingLevel.HEADING_2),
        ...String(group.narrative_diary || "").split("\n").map(line =>
          new Paragraph({ children: [new TextRun({ text: line, font: "Georgia", size: 22 })], spacing: { before: 40, after: 40 } })
        ),

        addHeading("Section 3: Key Events Timeline", HeadingLevel.HEADING_2),
        ...keyEvents.map(e => addBullet(typeof e === "object" ? JSON.stringify(e) : String(e))),

        addHeading("Section 4: People Directory", HeadingLevel.HEADING_2),
        ...peopleDir.map(p => addBullet(`${p["name"] || ""} — ${p["role"] || ""} (First: ${p["first_appearance"] || ""}, Mentions: ${p["mention_count"] || ""})`)),

        addHeading("Section 5: Places Directory", HeadingLevel.HEADING_2),
        ...placesDir.map(p => addBullet(`${p["place"] || ""} — ${p["context"] || ""} (Mentions: ${p["mention_count"] || ""})`)),

        addHeading("Section 6: Organizations", HeadingLevel.HEADING_2),
        ...orgsDir.map(o => addBullet(String(o))),

        addHeading("Section 7: Dominant Themes", HeadingLevel.HEADING_2),
        ...themes.map(t => addBullet(String(t))),

        addHeading("Section 8: Noteworthy Items", HeadingLevel.HEADING_2),
        ...noteworthy.map(n => addBullet(String(n))),

        addHeading("Section 9: Gaps & Unknowns", HeadingLevel.HEADING_2),
        addText(String(group.gaps_and_unknowns || "")),

        addHeading("Section 10: Bibliography", HeadingLevel.HEADING_2),
        ...bibliography.map(b => new Paragraph({ text: String(b), spacing: { before: 60, after: 60 } })),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  await fsp.writeFile(outPath, buffer);
}

async function generateMasterDocx(groups: Record<string, unknown>[], allItems: ProcessedItem[], outPath: string): Promise<void> {
  const addHeading = (text: string, level: HeadingLevel) =>
    new Paragraph({ text, heading: level, spacing: { before: 200, after: 100 } });
  const addText = (text: string, bold = false) =>
    new Paragraph({ children: [new TextRun({ text: text || "", bold, size: 22, font: "Georgia" })], spacing: { before: 60, after: 60 } });
  const addBullet = (text: string) =>
    new Paragraph({ text: `• ${text}`, indent: { left: 400 }, spacing: { before: 40, after: 40 } });

  const allPeople = new Map<string, number>();
  const allPlaces = new Map<string, number>();
  const allThemes: string[] = [];
  const allNoteworthy: string[] = [];
  const allBibliography = new Set<string>();
  const allKeyEvents: string[] = [];

  for (const g of groups) {
    if (Array.isArray(g.people_directory)) {
      for (const p of g.people_directory as Record<string, unknown>[]) {
        const name = String(p["name"] || "");
        allPeople.set(name, (allPeople.get(name) || 0) + Number(p["mention_count"] || 1));
      }
    }
    if (Array.isArray(g.places_directory)) {
      for (const p of g.places_directory as Record<string, unknown>[]) {
        const place = String(p["place"] || "");
        allPlaces.set(place, (allPlaces.get(place) || 0) + Number(p["mention_count"] || 1));
      }
    }
    if (Array.isArray(g.themes)) {
      allThemes.push(...(g.themes as string[]).map(String));
    }
    if (Array.isArray(g.noteworthy_items)) {
      allNoteworthy.push(...(g.noteworthy_items as string[]).map(String));
    }
    if (Array.isArray(g.bibliography)) {
      for (const b of g.bibliography as string[]) allBibliography.add(String(b));
    }
    if (Array.isArray(g.key_events)) {
      allKeyEvents.push(...(g.key_events as string[]).map(e => typeof e === "object" ? JSON.stringify(e) : String(e)));
    }
  }

  const itemTypes: Record<string, number> = {};
  let noteworthyCount = 0;
  for (const item of allItems) {
    const t = item.item_type || "other";
    itemTypes[t] = (itemTypes[t] || 0) + 1;
    if (item.noteworthy_flag) noteworthyCount++;
  }

  const statsTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total Items", bold: true })] })] }),
        new TableCell({ children: [new Paragraph(String(allItems.length))] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Item Types", bold: true })] })] }),
        new TableCell({ children: [new Paragraph(Object.entries(itemTypes).map(([k, v]) => `${k}: ${v}`).join(", "))] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total People Identified", bold: true })] })] }),
        new TableCell({ children: [new Paragraph(String(allPeople.size))] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total Places", bold: true })] })] }),
        new TableCell({ children: [new Paragraph(String(allPlaces.size))] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Noteworthy Items", bold: true })] })] }),
        new TableCell({ children: [new Paragraph(String(noteworthyCount))] }),
      ]}),
    ],
  });

  const doc = new Document({
    sections: [{
      children: [
        addHeading("LetterVault Pro — Complete Archive Analysis", HeadingLevel.HEADING_1),
        addText(`Total Items: ${allItems.length} | Groups: ${groups.length} | Generated: ${new Date().toISOString()}`, true),

        addHeading("Section 1: Collection Statistics", HeadingLevel.HEADING_2),
        statsTable,

        addHeading("Section 2: Group Summaries", HeadingLevel.HEADING_2),
        ...groups.flatMap(g => [
          addText(`${String(g.group_label)}: ${String(g.archivist_summary || "")}`, false),
          new Paragraph({ text: "", spacing: { before: 80, after: 80 } }),
        ]),

        addHeading("Section 3: Complete People Index", HeadingLevel.HEADING_2),
        ...[...allPeople.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([name, count]) => addBullet(`${name} (${count} mentions)`)),

        addHeading("Section 4: Complete Places Index", HeadingLevel.HEADING_2),
        ...[...allPlaces.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([place, count]) => addBullet(`${place} (${count} mentions)`)),

        addHeading("Section 5: Master Timeline", HeadingLevel.HEADING_2),
        ...allKeyEvents.map(e => addBullet(e)),

        addHeading("Section 6: Themes Across the Collection", HeadingLevel.HEADING_2),
        ...[...new Set(allThemes)].map(t => addBullet(t)),

        addHeading("Section 7: All Noteworthy Items", HeadingLevel.HEADING_2),
        ...[...new Set(allNoteworthy)].map(n => {
          const item = allItems.find(i => i.item_id === n);
          return addBullet(`${n}${item ? ` — ${item.noteworthy_reason || ""}` : ""}`);
        }),

        addHeading("Section 8: Master Bibliography", HeadingLevel.HEADING_2),
        ...[...allBibliography].sort().map(b => new Paragraph({ text: b, spacing: { before: 60, after: 60 } })),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  await fsp.writeFile(outPath, buffer);
}

router.get("/lettervault/status", (_req: Request, res: Response) => {
  const hasKey = !!process.env["ANTHROPIC_API_KEY"];
  res.json({
    items_count: itemStore.length,
    api_key: hasKey,
    default_item_schema: DEFAULT_ITEM_SCHEMA,
    default_analysis_schema: DEFAULT_ANALYSIS_SCHEMA,
    protocol: sessionStore.protocol || "",
    item_schema: sessionStore.itemSchema || DEFAULT_ITEM_SCHEMA,
    analysis_schema: sessionStore.analysisSchema || DEFAULT_ANALYSIS_SCHEMA,
  });
});

router.post("/lettervault/save-protocol", (req: Request, res: Response) => {
  const { protocol } = req.body as { protocol: string };
  if (!protocol || !protocol.trim()) {
    res.status(400).json({ error: "Protocol cannot be empty" });
    return;
  }
  sessionStore.protocol = protocol;
  res.json({ success: true, message: "Protocol saved successfully" });
});

router.post("/lettervault/save-schema", (req: Request, res: Response) => {
  const { schema } = req.body as { schema: string };
  try {
    JSON.parse(schema);
  } catch {
    res.status(400).json({ error: "Invalid JSON schema" });
    return;
  }
  sessionStore.itemSchema = schema;
  res.json({ success: true });
});

router.post("/lettervault/save-analysis-schema", (req: Request, res: Response) => {
  const { schema } = req.body as { schema: string };
  try {
    JSON.parse(schema);
  } catch {
    res.status(400).json({ error: "Invalid JSON schema" });
    return;
  }
  sessionStore.analysisSchema = schema;
  res.json({ success: true });
});

router.post("/lettervault/clear-session", (_req: Request, res: Response) => {
  itemStore.length = 0;
  sessionStore.protocol = undefined;
  sessionStore.itemSchema = DEFAULT_ITEM_SCHEMA;
  sessionStore.analysisSchema = DEFAULT_ANALYSIS_SCHEMA;
  try {
    fs.rmSync(ITEMS_DIR, { recursive: true, force: true });
    fs.rmSync(ANALYSIS_DIR, { recursive: true, force: true });
    fs.mkdirSync(ITEMS_DIR, { recursive: true });
    fs.mkdirSync(ANALYSIS_DIR, { recursive: true });
  } catch {
  }
  res.json({ success: true, message: "Session cleared" });
});

router.get("/lettervault/items", (_req: Request, res: Response) => {
  res.json({ items: itemStore });
});

router.post("/lettervault/process", upload.array("files", 50), async (req: Request, res: Response) => {
  const files = req.files as Express.Multer.File[];
  if (!files || files.length === 0) {
    res.status(400).json({ error: "No files uploaded" });
    return;
  }
  if (!sessionStore.protocol) {
    res.status(400).json({ error: "No protocol saved. Please save a protocol in Tab 1 first." });
    return;
  }

  let client: Anthropic;
  try {
    client = getAnthropicClient();
  } catch {
    res.status(500).json({ error: "ANTHROPIC_API_KEY is not configured. Please add it to Replit Secrets." });
    return;
  }

  const results: ProcessedItem[] = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (!file) continue;
    const itemId = uuidv4();
    const processedAt = new Date().toISOString();

    let imageData: string;
    let mediaType: "image/jpeg" | "image/png" | "image/gif" | "image/webp";

    try {
      const ext = path.extname(file.originalname).toLowerCase();
      if (ext === ".pdf") {
        results.push({
          item_id: itemId,
          filename: file.originalname,
          needs_review: true,
          raw_response: "PDF conversion not supported in this environment. Please convert to image first.",
          processed_at: processedAt,
        });
        continue;
      }
      imageData = file.buffer.toString("base64");
      mediaType = getMediaType(file.originalname);
    } catch (err) {
      results.push({
        item_id: itemId,
        filename: file.originalname,
        needs_review: true,
        raw_response: `Image conversion error: ${err}`,
        processed_at: processedAt,
      });
      continue;
    }

    const systemPrompt = `You are an expert archival document specialist processing historical items for a scholarly archive. You identify and transcribe letters, news clippings, photographs, postcards, typed documents, and handwritten materials. Follow all instructions precisely. Return ONLY valid JSON with no markdown, no code fences, no preamble.`;

    const userPrompt = `Item image is attached.

Processing Protocol to follow exactly:
${sessionStore.protocol}

Output Schema (fill every field):
${sessionStore.itemSchema || DEFAULT_ITEM_SCHEMA}

Instructions:
1. Identify the item_type from the image
2. Transcribe ALL visible text verbatim in full_transcription
3. For photographs or images with no text, describe what is depicted in full_transcription
4. Apply every step of the Processing Protocol
5. Extract all names, places, organizations into their arrays
6. Set noteworthy_flag to true if item has unusual historical significance
7. Write bibliography_entry in Chicago citation style
8. Set confidence_score based on image legibility
9. Return ONLY the JSON. No other text.`;

    let parsedItem: ProcessedItem | null = null;
    let attempts = 0;

    while (attempts < 2 && !parsedItem) {
      attempts++;
      try {
        const response = await client.messages.create({
          model: "claude-sonnet-4-5",
          max_tokens: 4000,
          system: systemPrompt,
          messages: [
            {
              role: "user",
              content: [
                {
                  type: "image",
                  source: {
                    type: "base64",
                    media_type: mediaType,
                    data: imageData,
                  },
                },
                {
                  type: "text",
                  text: attempts === 1 ? userPrompt : `${userPrompt}\n\nIMPORTANT: Your previous response was not valid JSON. Return ONLY raw JSON, no markdown fences, no explanation.`,
                },
              ],
            },
          ],
        });

        const rawText = response.content[0]?.type === "text" ? response.content[0].text : "";
        const cleaned = rawText.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "").trim();
        const parsed = JSON.parse(cleaned) as Record<string, unknown>;
        parsed["item_id"] = itemId;
        parsed["filename"] = file.originalname;
        parsed["processed_at"] = processedAt;
        parsedItem = parsed as unknown as ProcessedItem;
      } catch {
        if (attempts >= 2) {
          parsedItem = null;
        }
        await new Promise(r => setTimeout(r, 1000));
      }
    }

    if (!parsedItem) {
      const fallback: ProcessedItem = {
        item_id: itemId,
        filename: file.originalname,
        needs_review: true,
        raw_response: "Failed to parse Claude response after 2 attempts",
        processed_at: processedAt,
      };
      results.push(fallback);
      itemStore.push(fallback);
      continue;
    }

    const docxPath = path.join(ITEMS_DIR, `${itemId}.docx`);
    try {
      await generateItemDocx(parsedItem, docxPath);
      parsedItem.docx_path = docxPath;
    } catch (docxErr) {
      console.error("DOCX generation error:", docxErr);
    }

    itemStore.push(parsedItem);
    results.push(parsedItem);

    if (i < files.length - 1) {
      await new Promise(r => setTimeout(r, 1000));
    }
  }

  res.json({ success: true, processed: results.length, items: results });
});

router.post("/lettervault/analyze", async (req: Request, res: Response) => {
  if (itemStore.length < 1) {
    res.status(400).json({ error: "No items processed yet" });
    return;
  }

  const { grouping, customCriteria, depth } = req.body as {
    grouping: string;
    customCriteria?: string;
    depth?: string;
  };

  let client: Anthropic;
  try {
    client = getAnthropicClient();
  } catch {
    res.status(500).json({ error: "ANTHROPIC_API_KEY is not configured" });
    return;
  }

  const groups: Map<string, ProcessedItem[]> = new Map();

  const groupItem = (item: ProcessedItem): string => {
    switch (grouping) {
      case "year": {
        const d = item.date_on_item || "";
        const match = d.match(/\b(1[0-9]{3}|20[0-9]{2})\b/);
        return match ? match[1] : "Unknown Year";
      }
      case "decade": {
        const d = item.date_on_item || "";
        const match = d.match(/\b(1[0-9]{3}|20[0-9]{2})\b/);
        if (match) {
          const year = parseInt(match[1]);
          return `${Math.floor(year / 10) * 10}s`;
        }
        return "Unknown Decade";
      }
      case "sender":
        return item.sender_name || "Unknown Sender";
      case "recipient":
        return item.recipient_name || "Unknown Recipient";
      case "type":
        return item.item_type || "Unknown Type";
      case "era":
        return item.era_tag || "Unknown Era";
      case "noteworthy":
        return item.noteworthy_flag ? "Noteworthy Items" : "Standard Items";
      case "all":
        return "Complete Collection";
      default:
        return item.era_tag || item.item_type || "General";
    }
  };

  for (const item of itemStore) {
    const key = groupItem(item);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(item);
  }

  const systemPrompt = `You are a senior archivist and historian synthesizing a collection of historical documents for scholarly research. You write with authority, precision, and narrative clarity. Return ONLY valid JSON.`;

  const depthGuide = depth === "quick" ? "1-2 pages per group" : depth === "deep" ? "full synthesis, comprehensive detail" : "3-5 pages per group";

  const groupResults: Record<string, unknown>[] = [];

  for (const [label, items] of groups) {
    const userPrompt = `You are analyzing a group of ${items.length} archival items labeled: ${label}

Grouping criteria used: ${grouping}${customCriteria ? ` — Custom focus: ${customCriteria}` : ""}
Analysis depth: ${depthGuide}

Here is the complete data for all items in this group:
${JSON.stringify(items, null, 2)}

Analysis Schema to fill:
${sessionStore.analysisSchema || DEFAULT_ANALYSIS_SCHEMA}

Instructions:
1. Write narrative_diary as flowing chronological prose — synthesize the human story across all items as if writing a historical diary entry or chapter
2. Build complete people_directory, places_directory, organizations_directory from all items
3. Identify key_events with dates where determinable
4. Extract dominant themes
5. List all noteworthy_items by item_id
6. Compile full bibliography from all bibliography_entry fields
7. Write archivist_summary for a researcher encountering this group for the first time
8. Note gaps_and_unknowns — missing dates, unclear names, incomplete sequences
9. Return ONLY the JSON.`;

    let groupResult: Record<string, unknown> | null = null;
    let attempts = 0;

    while (attempts < 2 && !groupResult) {
      attempts++;
      try {
        const response = await client.messages.create({
          model: "claude-sonnet-4-5",
          max_tokens: 4000,
          system: systemPrompt,
          messages: [{ role: "user", content: userPrompt }],
        });
        const rawText = response.content[0]?.type === "text" ? response.content[0].text : "";
        const cleaned = rawText.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "").trim();
        const parsed = JSON.parse(cleaned) as Record<string, unknown>;
        parsed["group_label"] = label;
        parsed["item_count"] = items.length;
        groupResult = parsed;
      } catch {
        await new Promise(r => setTimeout(r, 2000));
      }
    }

    if (!groupResult) {
      groupResult = {
        group_label: label,
        item_count: items.length,
        archivist_summary: "Analysis failed — Claude response could not be parsed",
        narrative_diary: "",
        key_events: [],
        people_directory: [],
        places_directory: [],
        organizations_directory: [],
        themes: [],
        noteworthy_items: [],
        bibliography: [],
        gaps_and_unknowns: "Analysis failed",
        date_range: "",
        group_criteria: grouping,
      };
    }

    const groupDocxPath = path.join(ANALYSIS_DIR, `${label.replace(/[^a-z0-9]/gi, "_")}.docx`);
    try {
      await generateGroupDocx(groupResult, groupDocxPath);
    } catch (err) {
      console.error("Group docx error:", err);
    }

    groupResults.push(groupResult);
  }

  const masterPath = path.join(ANALYSIS_DIR, "MASTER_REPORT.docx");
  try {
    await generateMasterDocx(groupResults, itemStore, masterPath);
  } catch (err) {
    console.error("Master report error:", err);
  }

  res.json({ success: true, groups: groupResults, master_report: "MASTER_REPORT.docx" });
});

router.get("/lettervault/download/item/:itemId", (req: Request, res: Response) => {
  const { itemId } = req.params;
  const filePath = path.join(ITEMS_DIR, `${itemId}.docx`);
  if (!fs.existsSync(filePath)) {
    res.status(404).json({ error: "File not found" });
    return;
  }
  res.download(filePath, `item_${itemId}.docx`);
});

router.get("/lettervault/download/analysis/:filename", (req: Request, res: Response) => {
  const { filename } = req.params;
  const safe = filename.replace(/\.\./g, "");
  const filePath = path.join(ANALYSIS_DIR, safe);
  if (!fs.existsSync(filePath)) {
    res.status(404).json({ error: "File not found" });
    return;
  }
  res.download(filePath, safe);
});

router.get("/lettervault/download-zip/items", (_req: Request, res: Response) => {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", "attachment; filename=lettervault_items.zip");
  const archive = new ZipArchive();
  archive.pipe(res);
  archive.directory(ITEMS_DIR, "items");
  archive.finalize().catch(err => { console.error(err); });
});

router.get("/lettervault/download-zip/analysis", (_req: Request, res: Response) => {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", "attachment; filename=lettervault_analysis.zip");
  const archive = new ZipArchive();
  archive.pipe(res);
  archive.directory(ANALYSIS_DIR, "analysis");
  archive.finalize().catch(err => { console.error(err); });
});

router.get("/lettervault/download-zip/all", (_req: Request, res: Response) => {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", "attachment; filename=lettervault_complete_archive.zip");
  const archive = new ZipArchive();
  archive.pipe(res);
  archive.directory(ITEMS_DIR, "items");
  archive.directory(ANALYSIS_DIR, "analysis");
  archive.finalize().catch(err => { console.error(err); });
});

router.get("/lettervault/analysis-files", (_req: Request, res: Response) => {
  try {
    const files = fs.readdirSync(ANALYSIS_DIR).filter(f => f.endsWith(".docx"));
    res.json({ files });
  } catch {
    res.json({ files: [] });
  }
});

export default router;
