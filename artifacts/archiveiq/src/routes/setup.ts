import { Router } from "express";
import { BUS, getDefaultItemSchema, getDefaultAnalysisSchema } from "../lib/intelligence-bus.js";

const router = Router();

const TEMPLATES: Record<string, { item: string; analysis: string; protocol: string }> = {
  "Archival Diary Extraction": {
    protocol: "ARCHIVAL DIARY EXTRACTION PROTOCOL\n\nFor each diary page:\n1. Identify the date header\n2. Transcribe all text verbatim\n3. Extract all persons, places, events\n4. Flag any OCR errors or illegible passages\n5. Note any significant historical events mentioned\n6. Assign confidence score based on legibility",
    item: JSON.stringify({ item_id: "UUID", date: "string", transcription: "string", people: [], places: [], events: [], ocr_notes: "string", confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", narrative: "string", persons_directory: [], places_directory: [], themes: [], gaps: "string" }, null, 2),
  },
  "Correspondence Indexing": {
    protocol: "CORRESPONDENCE INDEXING PROTOCOL\n\n1. Identify sender and recipient\n2. Extract date and location\n3. Transcribe full text\n4. Note all persons mentioned\n5. Identify key topics\n6. Flag items of unusual significance",
    item: JSON.stringify({ item_id: "UUID", date: "string", sender: "string", recipient: "string", location: "string", transcription: "string", people_mentioned: [], topics: [], significance_flag: false }, null, 2),
    analysis: JSON.stringify({ group_label: "string", narrative: "string", correspondence_network: [], key_topics: [], timeline: [], bibliography: [] }, null, 2),
  },
  "News Clipping Archive": {
    protocol: "NEWS CLIPPING ARCHIVE PROTOCOL\n\n1. Identify publication, date, and headline\n2. Transcribe article text\n3. Extract named persons and places\n4. Identify primary topics\n5. Note geographic scope",
    item: JSON.stringify({ item_id: "UUID", publication: "string", date: "string", headline: "string", transcription: "string", persons: [], places: [], topics: [], confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", thematic_summary: "string", persons_index: [], places_index: [], topic_breakdown: [] }, null, 2),
  },
  "Photograph Cataloguing": {
    protocol: "PHOTOGRAPH CATALOGUING PROTOCOL\n\n1. Describe visual content\n2. Identify persons if possible\n3. Estimate date range from visual evidence\n4. Note location if identifiable\n5. Describe condition and any inscriptions",
    item: JSON.stringify({ item_id: "UUID", description: "string", persons_identified: [], estimated_date: "string", location: "string", inscriptions: "string", condition: "string", confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", visual_narrative: "string", persons_index: [], locations_index: [], date_range: "string" }, null, 2),
  },
  "Legal Document Review": {
    protocol: "LEGAL DOCUMENT REVIEW PROTOCOL\n\n1. Identify document type\n2. List all parties\n3. Extract key dates\n4. Summarize legal terms and obligations\n5. Note property or monetary values\n6. Flag unusual clauses",
    item: JSON.stringify({ item_id: "UUID", document_type: "string", parties: [], date: "string", jurisdiction: "string", summary: "string", key_terms: [], monetary_values: [], property_references: [], confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", legal_narrative: "string", parties_index: [], property_index: [], timeline: [] }, null, 2),
  },
  "Military Records": {
    protocol: "MILITARY RECORDS PROTOCOL\n\n1. Identify service member\n2. Extract rank, unit, and dates of service\n3. Note battles, campaigns, or postings\n4. Record any awards or citations\n5. Note discharge type and date",
    item: JSON.stringify({ item_id: "UUID", service_member: "string", rank: "string", unit: "string", service_dates: "string", campaigns: [], awards: [], discharge_type: "string", confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", unit_narrative: "string", personnel_index: [], battles_index: [], timeline: [] }, null, 2),
  },
  "Church Records": {
    protocol: "CHURCH RECORDS PROTOCOL\n\n1. Identify record type (baptism/marriage/burial/membership)\n2. Extract all persons named\n3. Record dates and officiants\n4. Note sponsors or witnesses\n5. Extract any genealogical information",
    item: JSON.stringify({ item_id: "UUID", record_type: "string", persons: [], date: "string", officiant: "string", witnesses: [], notes: "string", confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", congregation_narrative: "string", family_networks: [], timeline: [], persons_index: [] }, null, 2),
  },
  "Custom": {
    protocol: "",
    item: JSON.stringify({ item_id: "UUID", content: "string", metadata: {}, confidence_score: "high|medium|low" }, null, 2),
    analysis: JSON.stringify({ group_label: "string", narrative: "string", index: [] }, null, 2),
  },
};

router.get("/:projectId", (req, res) => {
  const { projectId } = req.params;
  const protocols = BUS.protocolRegistry[projectId] ?? [];
  const schemas = BUS.schemaRegistry[projectId] ?? [];
  const latestProtocol = protocols[protocols.length - 1]?.content ?? "";
  const latestSchema = schemas[schemas.length - 1];

  res.json({
    protocol: latestProtocol,
    item_schema: latestSchema?.item_schema ?? getDefaultItemSchema(BUS.projects[projectId]?.type ?? ""),
    analysis_schema: latestSchema?.analysis_schema ?? getDefaultAnalysisSchema(BUS.projects[projectId]?.type ?? ""),
    protocol_version: protocols.length,
    schema_version: schemas.length,
    templates: Object.keys(TEMPLATES),
  });
});

router.get("/history/:projectId", (req, res) => {
  const { projectId } = req.params;
  res.json({
    protocols: (BUS.protocolRegistry[projectId] ?? []).map(v => ({
      version: v.version,
      saved_at: v.saved_at,
      preview: v.content.substring(0, 200) + "...",
    })),
    schemas: (BUS.schemaRegistry[projectId] ?? []).map(v => ({
      version: v.version,
      saved_at: v.saved_at,
    })),
  });
});

router.post("/protocol", (req, res) => {
  const { projectId, content, restore_version } = req.body;
  if (!projectId) return res.status(400).json({ error: "projectId required" });

  if (!BUS.protocolRegistry[projectId]) BUS.protocolRegistry[projectId] = [];

  let finalContent = content;
  if (restore_version != null) {
    const v = BUS.protocolRegistry[projectId][restore_version - 1];
    if (!v) return res.status(404).json({ error: "Version not found" });
    finalContent = v.content;
  }

  const nextVersion = (BUS.protocolRegistry[projectId].length) + 1;
  BUS.protocolRegistry[projectId].push({ version: nextVersion, content: finalContent, saved_at: new Date().toISOString() });
  BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "protocol_saved", project_id: projectId });
  res.json({ success: true, version: nextVersion });
});

router.post("/schema", (req, res) => {
  const { projectId, item_schema, analysis_schema } = req.body;
  if (!projectId) return res.status(400).json({ error: "projectId required" });

  if (!BUS.schemaRegistry[projectId]) BUS.schemaRegistry[projectId] = [];
  const nextVersion = (BUS.schemaRegistry[projectId].length) + 1;
  const proj = BUS.projects[projectId];

  BUS.schemaRegistry[projectId].push({
    version: nextVersion,
    item_schema: item_schema ?? getDefaultItemSchema(proj?.type ?? ""),
    analysis_schema: analysis_schema ?? getDefaultAnalysisSchema(proj?.type ?? ""),
    saved_at: new Date().toISOString(),
  });
  BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "schema_saved", project_id: projectId });
  res.json({ success: true, version: nextVersion });
});

router.post("/template", (req, res) => {
  const { templateName } = req.body;
  const t = TEMPLATES[templateName];
  if (!t) return res.status(404).json({ error: "Template not found" });
  res.json(t);
});

export default router;
