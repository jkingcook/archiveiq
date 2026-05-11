import { Router } from "express";
import Anthropic from "@anthropic-ai/sdk";
import { BUS, getActiveProject } from "../lib/intelligence-bus.js";
import type { MachensItem } from "../lib/intelligence-bus.js";

const router = Router();
const client = new Anthropic({ apiKey: process.env["ANTHROPIC_API_KEY"] });

router.get("/patterns", (_req, res) => {
  const topPeople = Object.entries(BUS.patternLibrary.people)
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 10)
    .map(([name, data]) => ({ name, count: data.count, projects: Array.from(data.projects) }));

  const topPlaces = Object.entries(BUS.patternLibrary.places)
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 10)
    .map(([name, data]) => ({ name, count: data.count, projects: Array.from(data.projects) }));

  const dateBreakdown = Object.entries(BUS.patternLibrary.dates)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([year, data]) => ({ year, count: data.count }));

  const categoryBreakdown = Object.entries(BUS.patternLibrary.categories)
    .sort((a, b) => b[1].count - a[1].count)
    .map(([cat, data]) => ({ category: cat, count: data.count }));

  const confidenceDist = { high: 0, medium: 0, low: 0 };
  const noteworthy: string[] = [];
  for (const item of BUS.itemStore) {
    const conf = (item as { confidence_score?: string }).confidence_score ?? "low";
    if (conf in confidenceDist) confidenceDist[conf as keyof typeof confidenceDist]++;
    if ((item as { noteworthy_flag?: boolean }).noteworthy_flag) {
      noteworthy.push(item.item_id);
    }
  }

  const docTypeBreakdown: Record<string, number> = {};
  for (const item of BUS.itemStore) {
    const t = (item as MachensItem).source_type ?? (item as { item_type?: string }).item_type ?? "unknown";
    docTypeBreakdown[t] = (docTypeBreakdown[t] ?? 0) + 1;
  }

  res.json({
    totalItems: BUS.itemStore.length,
    totalRows: BUS.itemStore.reduce((s, i) => s + ((i as MachensItem).register_rows?.length ?? 0), 0),
    topPeople,
    topPlaces,
    dateBreakdown,
    categoryBreakdown,
    confidenceDist,
    noteworthyCount: noteworthy.length,
    noteworthyItems: noteworthy,
    docTypeBreakdown,
  });
});

router.get("/person/:name", (req, res) => {
  const searchName = decodeURIComponent(req.params.name).toLowerCase();
  const matchingItems = BUS.itemStore.filter(item => {
    const people = (item as MachensItem).people_extracted ?? (item as { people_mentioned?: string[] }).people_mentioned ?? [];
    return people.some(p => p.toLowerCase().includes(searchName));
  });

  const patternEntry = Object.entries(BUS.patternLibrary.people).find(([k]) => k.includes(searchName));
  const voiceRows: MachensItem["register_rows"] = [];
  const allMentions: Array<{ item_id: string; project_id: string; date?: string; context: string }> = [];

  for (const item of matchingItems) {
    const rows = (item as MachensItem).register_rows ?? [];
    for (const row of rows) {
      if (row.person?.toLowerCase().includes(searchName) || row.fact?.toLowerCase().includes(searchName)) {
        allMentions.push({ item_id: item.item_id, project_id: item.project_id, date: row.date, context: row.fact });
        if (row.category === "Voice") voiceRows.push(row);
      }
    }
    if (rows.length === 0) {
      allMentions.push({ item_id: item.item_id, project_id: item.project_id, context: "Found in item" });
    }
  }

  const dateRange = allMentions
    .map(m => m.date)
    .filter(Boolean)
    .sort();

  res.json({
    searchName,
    itemCount: matchingItems.length,
    projects: patternEntry ? Array.from(patternEntry[1].projects) : [],
    dateRange: dateRange.length ? { from: dateRange[0], to: dateRange[dateRange.length - 1] } : null,
    mentions: allMentions,
    voiceRows,
    crossProjectConnections: Object.entries(BUS.crossProjectIndex.sharedPeople).find(([k]) => k.includes(searchName))?.[1] ?? [],
  });
});

router.get("/place/:name", (req, res) => {
  const searchName = decodeURIComponent(req.params.name).toLowerCase();
  const matchingItems = BUS.itemStore.filter(item => {
    const places = (item as MachensItem).places_extracted ?? (item as { places_mentioned?: string[] }).places_mentioned ?? [];
    return places.some(p => p.toLowerCase().includes(searchName));
  });

  const allMentions: Array<{ item_id: string; project_id: string; date?: string; context: string }> = [];
  for (const item of matchingItems) {
    const rows = (item as MachensItem).register_rows ?? [];
    for (const row of rows) {
      if (row.fact?.toLowerCase().includes(searchName)) {
        allMentions.push({ item_id: item.item_id, project_id: item.project_id, date: row.date, context: row.fact });
      }
    }
    if (rows.length === 0) {
      allMentions.push({ item_id: item.item_id, project_id: item.project_id, context: "Found in item" });
    }
  }

  res.json({
    searchName,
    itemCount: matchingItems.length,
    mentions: allMentions,
    crossProjectConnections: Object.entries(BUS.crossProjectIndex.sharedPlaces).find(([k]) => k.includes(searchName))?.[1] ?? [],
  });
});

router.get("/cross-project", (_req, res) => {
  res.json({
    sharedPeople: Object.entries(BUS.crossProjectIndex.sharedPeople).map(([name, projects]) => ({ name, projects })),
    sharedPlaces: Object.entries(BUS.crossProjectIndex.sharedPlaces).map(([name, projects]) => ({ name, projects })),
    projectCount: Object.keys(BUS.projects).length,
    totalItems: BUS.itemStore.length,
  });
});

router.post("/ask", async (req, res) => {
  const { question, projectId } = req.body;
  if (!question) return res.status(400).json({ error: "question required" });
  if (!process.env["ANTHROPIC_API_KEY"]) return res.status(400).json({ error: "API key not set" });

  const scopedItems = projectId
    ? BUS.itemStore.filter(i => i.project_id === projectId)
    : BUS.itemStore;

  const busSummary = {
    projects: Object.values(BUS.projects).map(p => ({ id: p.id, name: p.name, type: p.type })),
    totalItems: BUS.itemStore.length,
    topPeople: Object.entries(BUS.patternLibrary.people).sort((a, b) => b[1].count - a[1].count).slice(0, 20).map(([n, d]) => ({ name: n, count: d.count })),
    topPlaces: Object.entries(BUS.patternLibrary.places).sort((a, b) => b[1].count - a[1].count).slice(0, 20).map(([n, d]) => ({ name: n, count: d.count })),
    items: scopedItems.slice(0, 20).map(i => ({
      id: i.item_id,
      project: i.project_id,
      year: (i as MachensItem).source_year ?? (i as { date_on_item?: string }).date_on_item,
      rows: (i as MachensItem).register_rows?.slice(0, 10),
      noteworthy: (i as MachensItem).noteworthy_flag,
    })),
  };

  const response = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 2000,
    system: `You are a senior archivist and historian with full knowledge of the ArchiveIQ collection. Answer the user's question based on the provided collection data. Cite specific item_ids when relevant. Be scholarly but accessible.`,
    messages: [{
      role: "user",
      content: `Question: ${question}\n\nCollection data:\n${JSON.stringify(busSummary, null, 2).substring(0, 6000)}`,
    }],
  });

  const answer = response.content[0].type === "text" ? response.content[0].text : "";
  BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "intelligence_ask", details: question.substring(0, 100) });
  res.json({ question, answer });
});

export default router;
