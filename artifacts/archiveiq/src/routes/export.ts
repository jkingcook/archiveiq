import { Router } from "express";
import Anthropic from "@anthropic-ai/sdk";
import { createReadStream, existsSync, readdirSync, writeFileSync } from "fs";
import { join } from "path";
import { BUS, getProjectItems, getProjectAnalysis } from "../lib/intelligence-bus.js";
import type { MachensItem } from "../lib/intelligence-bus.js";

const router = Router();
const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

router.get("/item/:id", (req, res) => {
  const item = BUS.itemStore.find(i => i.item_id === req.params.id);
  if (!item) return res.status(404).json({ error: "Not found" });
  const docxPath = (item as { docx_path?: string }).docx_path;
  if (!docxPath || !existsSync(docxPath)) {
    return res.status(404).json({ error: "Docx not generated yet" });
  }
  const filename = docxPath.split("/").pop() ?? "item.docx";
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  createReadStream(docxPath).pipe(res);
});

router.get("/analysis/:id", (req, res) => {
  const analysis = BUS.analysisStore.find(a => a.id === req.params.id);
  if (!analysis) return res.status(404).json({ error: "Not found" });
  const docxPath = analysis.docx_path;
  if (!docxPath || !existsSync(docxPath)) {
    return res.status(404).json({ error: "Docx not generated" });
  }
  const filename = docxPath.split("/").pop() ?? "analysis.docx";
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  createReadStream(docxPath).pipe(res);
});

router.get("/master/:projectId", (req, res) => {
  const { projectId } = req.params;
  const project = BUS.projects[projectId];
  if (!project) return res.status(404).json({ error: "Project not found" });

  const files = existsSync("/tmp/output/analysis") ? readdirSync("/tmp/output/analysis") : [];
  const masterFile = files.find(f => f.includes("Master") || f.includes("master"));
  if (!masterFile) return res.status(404).json({ error: "Master report not generated yet — run analysis first" });

  const fullPath = `/tmp/output/analysis/${masterFile}`;
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", `attachment; filename="${masterFile}"`);
  createReadStream(fullPath).pipe(res);
});

router.get("/zip/:type", async (req, res) => {
  const { type } = req.params;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const archiverMod = (await import("archiver")) as any;
  const archive = new archiverMod.ZipArchive();
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="archiveiq_${type}_${Date.now()}.zip"`);
  archive.pipe(res);

  if (type === "items" || type === "all") {
    if (existsSync("/tmp/output/items")) {
      archive.directory("/tmp/output/items", "items");
    }
  }
  if (type === "analysis" || type === "all") {
    if (existsSync("/tmp/output/analysis")) {
      archive.directory("/tmp/output/analysis", "analysis");
    }
  }

  archive.finalize();
});

router.get("/csv/:type", (req, res) => {
  const { type } = req.params;
  const { projectId } = req.query;

  let csvContent = "";

  if (type === "rows") {
    const items = projectId
      ? BUS.itemStore.filter(i => i.project_id === String(projectId))
      : BUS.itemStore;

    const headers = ["item_id", "project_id", "source_year", "date", "person", "category", "fact", "source"];
    const rows: string[][] = [headers];

    for (const item of items) {
      const m = item as MachensItem;
      for (const row of m.register_rows ?? []) {
        rows.push([
          item.item_id,
          item.project_id,
          m.source_year ?? "",
          row.date ?? "",
          row.person ?? "",
          row.category ?? "",
          `"${(row.fact ?? "").replace(/"/g, '""')}"`,
          `"${(row.source ?? "").replace(/"/g, '""')}"`,
        ]);
      }
    }
    csvContent = rows.map(r => r.join(",")).join("\n");
    res.setHeader("Content-Disposition", 'attachment; filename="archiveiq_rows.csv"');

  } else if (type === "timeline") {
    const headers = ["date", "year", "person", "category", "fact", "item_id", "project_id"];
    const rows: string[][] = [headers];
    const items = projectId
      ? BUS.itemStore.filter(i => i.project_id === String(projectId))
      : BUS.itemStore;

    for (const item of items) {
      const m = item as MachensItem;
      for (const row of m.register_rows ?? []) {
        if (row.date) {
          rows.push([row.date, m.source_year ?? "", row.person ?? "", row.category ?? "", `"${(row.fact ?? "").replace(/"/g, '""')}"`, item.item_id, item.project_id]);
        }
      }
    }
    csvContent = rows.map(r => r.join(",")).join("\n");
    res.setHeader("Content-Disposition", 'attachment; filename="archiveiq_timeline.csv"');

  } else {
    const headers = ["item_id", "project_id", "source_year", "source_type", "confidence_score", "noteworthy", "row_count"];
    const rows: string[][] = [headers];
    const items = projectId
      ? BUS.itemStore.filter(i => i.project_id === String(projectId))
      : BUS.itemStore;

    for (const item of items) {
      const m = item as MachensItem;
      rows.push([item.item_id, item.project_id, m.source_year ?? "", m.source_type ?? (item as { item_type?: string }).item_type ?? "", m.confidence_score ?? "", String(m.noteworthy_flag ?? false), String(m.register_rows?.length ?? 0)]);
    }
    csvContent = rows.map(r => r.join(",")).join("\n");
    res.setHeader("Content-Disposition", 'attachment; filename="archiveiq_items.csv"');
  }

  res.setHeader("Content-Type", "text/csv");
  res.send(csvContent);
});

router.post("/custom", async (req, res) => {
  const { description, projectId } = req.body;
  if (!description) return res.status(400).json({ error: "description required" });
  if (!process.env["ANTHROPIC_API_KEY"]) return res.status(400).json({ error: "API key not set" });

  const items = projectId
    ? BUS.itemStore.filter(i => i.project_id === String(projectId))
    : BUS.itemStore;

  const response = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 4000,
    system: "You are ArchiveIQ's export builder. Build custom exports from the intelligence bus. Return a JSON object with { format: 'csv'|'json', filename: string, content: string }. The content field should contain the complete CSV or JSON string.",
    messages: [{
      role: "user",
      content: `Build this export: ${description}\n\nAvailable items (${items.length} total):\n${JSON.stringify(items.slice(0, 15), null, 2).substring(0, 5000)}`,
    }],
  });

  const text = response.content[0].type === "text" ? response.content[0].text : "";
  try {
    const cleaned = text.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
    const result = JSON.parse(cleaned) as { format: string; filename: string; content: string };
    const outPath = `/tmp/output/${result.filename ?? "custom_export"}`;
    writeFileSync(outPath, result.content);
    res.setHeader("Content-Disposition", `attachment; filename="${result.filename ?? "export"}"`);
    res.setHeader("Content-Type", result.format === "csv" ? "text/csv" : "application/json");
    res.send(result.content);
  } catch (_e) {
    res.json({ success: true, raw: text });
  }
});

router.get("/list", (_req, res) => {
  const itemFiles = existsSync("/tmp/output/items") ? readdirSync("/tmp/output/items").map(f => ({ name: f, path: f, type: "item" })) : [];
  const analysisFiles = existsSync("/tmp/output/analysis") ? readdirSync("/tmp/output/analysis").map(f => ({ name: f, path: f, type: "analysis" })) : [];
  res.json({ items: itemFiles, analysis: analysisFiles });
});

export default router;
