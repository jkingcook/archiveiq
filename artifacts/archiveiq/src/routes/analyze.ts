import { Router } from "express";
import Anthropic from "@anthropic-ai/sdk";
import { v4 as uuidv4 } from "uuid";
import { BUS, getProjectItems, getProjectAnalysis, isMachenProject } from "../lib/intelligence-bus.js";
import { buildGroupAnalysisDocx, buildMasterReportDocx } from "../lib/docx-builder.js";
import type { AnalysisResult, MachensItem } from "../lib/intelligence-bus.js";

const router = Router();
const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }

async function runGroupAnalysis(
  projectId: string,
  groupLabel: string,
  criteria: string,
  groupItems: MachensItem[],
  depth: string
): Promise<AnalysisResult> {
  const project = BUS.projects[projectId];
  const schemas = BUS.schemaRegistry[projectId] ?? [];
  const analysisSchema = schemas[schemas.length - 1]?.analysis_schema ?? "{}";
  const isMachen = isMachenProject(projectId);

  const maxTokens = depth === "deep" ? 4000 : depth === "standard" ? 3000 : 2000;

  const systemPrompt = `You are ArchiveIQ's collection intelligence engine analyzing materials from PROJECT: ${project.name}. You synthesize across multiple processed documents to produce authoritative narrative, directory, and reference outputs. You write the narrative_diary / narrative_summary as flowing prose at the level of a published scholarly work. Return ONLY valid JSON.`;

  const totalRows = groupItems.reduce((s, i) => s + (i.register_rows?.length ?? 0), 0);

  const userText = `Analyzing ${groupItems.length} items from ${project.name}, group: ${groupLabel}
Criteria: ${criteria}
Total register rows: ${totalRows}

All item data (summarized):
${JSON.stringify(groupItems.map(i => ({
    item_id: i.item_id,
    source_year: i.source_year,
    source_type: i.source_type,
    register_rows: i.register_rows?.slice(0, 30),
    people_extracted: i.people_extracted?.slice(0, 20),
    places_extracted: i.places_extracted?.slice(0, 20),
    dates_extracted: i.dates_extracted?.slice(0, 20),
    noteworthy_flag: i.noteworthy_flag,
    noteworthy_reason: i.noteworthy_reason,
    pass_1_coverage_report: i.pass_1_coverage_report,
  })), null, 2).substring(0, 8000)}

Analysis Schema:
${analysisSchema}

Instructions:
1. Write narrative prose synthesizing all items chronologically — tell the human story across the documents at scholarly level
2. Build complete persons directory with roles and priority flags
3. Build places directory
${isMachen ? `4. Compile donee_summary row counts; identify all Voice rows; note cross-reference opportunities` : `4. Identify dominant themes; compile bibliography`}
5. Write archivist_summary suitable for presenting to an institution
6. Note all gaps, unknowns, and unresolved questions
7. Return ONLY the JSON.`;

  let parsed: Record<string, unknown> | null = null;
  let rawText = "";

  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const response = await client.messages.create({
        model: "claude-sonnet-4-5",
        max_tokens: maxTokens,
        system: systemPrompt,
        messages: [{ role: "user", content: userText }],
      });
      rawText = response.content[0].type === "text" ? response.content[0].text : "";
      const cleaned = rawText.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
      parsed = JSON.parse(cleaned);
      break;
    } catch (_e) {
      if (attempt === 0) await sleep(2000);
    }
  }

  const analysisId = uuidv4();
  const now = new Date().toISOString();

  const result: AnalysisResult = {
    id: analysisId,
    project_id: projectId,
    group_label: groupLabel,
    criteria,
    item_count: groupItems.length,
    created_at: now,
    data: parsed ?? { error: "Parse failed", raw: rawText.substring(0, 500) },
  };

  try {
    const safeName = groupLabel.replace(/[^a-zA-Z0-9._-]/g, "_");
    let outPath: string;
    if (isMachen) {
      outPath = `/tmp/output/analysis/Biographical_Facts_Register_${safeName}.docx`;
    } else {
      outPath = `/tmp/output/analysis/${project.name.replace(/\s+/g, "_")}_${safeName}_Analysis.docx`;
    }
    await buildGroupAnalysisDocx(result, project, outPath);
    result.docx_path = outPath;
  } catch (_e) {
    // docx generation is non-fatal
  }

  BUS.analysisStore.push(result);
  BUS.sessionLog.push({ timestamp: now, action: "analysis_complete", project_id: projectId, details: groupLabel });

  return result;
}

router.post("/", async (req, res) => {
  const { projectId, groupBy, customCriteria, depth } = req.body;
  if (!projectId || !BUS.projects[projectId]) {
    return res.status(400).json({ error: "Valid projectId required" });
  }

  if (!process.env["ANTHROPIC_API_KEY"]) return res.status(400).json({ error: "ANTHROPIC_API_KEY not set" });

  const project = BUS.projects[projectId];
  const allItems = getProjectItems(projectId) as MachensItem[];
  if (allItems.length === 0) return res.status(400).json({ error: "No items to analyze" });

  const isMachen = isMachenProject(projectId);
  const results: AnalysisResult[] = [];

  function groupItems(): Map<string, MachensItem[]> {
    const groups = new Map<string, MachensItem[]>();

    if (customCriteria) {
      groups.set(customCriteria, allItems);
    } else if (groupBy === "all" || !groupBy) {
      groups.set("All Items", allItems);
    } else if (groupBy === "year") {
      for (const item of allItems) {
        const y = item.source_year || "Unknown Year";
        if (!groups.has(y)) groups.set(y, []);
        groups.get(y)!.push(item);
      }
    } else if (groupBy === "donee_priority" && isMachen) {
      const cats = ["A", "B", "C", "D", "E"];
      for (const cat of cats) {
        const filtered = allItems.filter(i => (i.donee_priority_categories ?? []).includes(cat));
        if (filtered.length > 0) groups.set(`Donee Category ${cat}`, filtered);
      }
    } else if (groupBy === "person") {
      const personMap = new Map<string, MachensItem[]>();
      for (const item of allItems) {
        for (const p of item.people_extracted ?? []) {
          if (!personMap.has(p)) personMap.set(p, []);
          personMap.get(p)!.push(item);
        }
      }
      for (const [p, items] of personMap) {
        if (items.length >= 2) groups.set(`Person: ${p}`, items);
      }
      if (groups.size === 0) groups.set("All Items", allItems);
    } else if (groupBy === "tier" && isMachen) {
      const tier1 = ["1899","1900","1901","1913","1925","1926","1927","1928"];
      const tier2 = ["1894","1895","1896","1897","1898","1914","1915","1916","1917","1918","1919","1920","1921","1922","1923","1924"];
      const g1 = allItems.filter(i => tier1.includes(i.source_year ?? ""));
      const g2 = allItems.filter(i => tier2.includes(i.source_year ?? ""));
      const g3 = allItems.filter(i => !tier1.includes(i.source_year ?? "") && !tier2.includes(i.source_year ?? ""));
      if (g1.length > 0) groups.set("Tier 1 Priority", g1);
      if (g2.length > 0) groups.set("Tier 2 Priority", g2);
      if (g3.length > 0) groups.set("Tier 3 / Other", g3);
    } else if (groupBy === "noteworthy") {
      const nw = allItems.filter(i => i.noteworthy_flag);
      if (nw.length > 0) groups.set("Noteworthy Items", nw);
      else groups.set("All Items (none noteworthy)", allItems);
    } else if (groupBy === "category") {
      const cats = new Set<string>();
      for (const item of allItems) {
        for (const row of item.register_rows ?? []) cats.add(row.category);
      }
      for (const cat of cats) {
        const filtered = allItems.filter(i => (i.register_rows ?? []).some(r => r.category === cat));
        if (filtered.length > 0) groups.set(`Category: ${cat}`, filtered);
      }
    } else {
      groups.set("All Items", allItems);
    }
    return groups;
  }

  const groups = groupItems();

  for (const [label, items] of groups) {
    try {
      const result = await runGroupAnalysis(projectId, label, customCriteria ?? groupBy ?? "all items", items, depth ?? "standard");
      results.push(result);
    } catch (e) {
      BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "analysis_error", project_id: projectId, details: String(e) });
    }
    await sleep(1000);
  }

  // Generate master report
  try {
    const now = new Date().toISOString().split("T")[0];
    let masterPath: string;
    if (isMachen) {
      masterPath = `/tmp/output/analysis/Machen_Family_Papers_Master_Register_${now}.docx`;
    } else {
      masterPath = `/tmp/output/analysis/${project.name.replace(/\s+/g, "_")}_Master_Report_${now}.docx`;
    }
    await buildMasterReportDocx(projectId, project, allItems, results, masterPath);
    BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "master_report_generated", project_id: projectId, details: masterPath });
  } catch (_e) {
    // master report is non-fatal
  }

  res.json({ success: true, groups: results.length, results });
});

router.get("/", (_req, res) => {
  res.json({ analysis: BUS.analysisStore });
});

router.get("/:projectId", (req, res) => {
  res.json({ analysis: getProjectAnalysis(req.params.projectId) });
});

export default router;
