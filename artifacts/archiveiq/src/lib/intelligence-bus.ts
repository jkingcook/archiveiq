import { v4 as uuidv4 } from "uuid";

export interface RegisterRow {
  date: string;
  person: string;
  category: string;
  fact: string;
  source: string;
}

export interface MachensItem {
  item_id: string;
  project_id: string;
  source_year: string;
  source_type: string;
  pass_applied: string;
  register_rows: RegisterRow[];
  donee_priority_categories: string[];
  headline_source_flag: boolean;
  pass_1_coverage_report: string | null;
  pass_2_anomalies: {
    category_1_ocr_errors: string[];
    category_2_multiword_confusion: string[];
    category_3_name_ambiguities: string[];
    category_4_lacunae: string[];
    category_6_period_usages_preserved: string[];
    category_7_substantive_flags: string[];
  };
  noteworthy_flag: boolean;
  noteworthy_reason: string | null;
  strengthens_case: string[];
  people_extracted: string[];
  places_extracted: string[];
  dates_extracted: string[];
  confidence_score: string;
  ocr_correction_notes: string;
  raw_text?: string;
  needs_review?: boolean;
  processed_at?: string;
  filename?: string;
  status?: string;
  docx_path?: string;
}

export interface GenericItem {
  item_id: string;
  project_id: string;
  item_type: string;
  date_on_item: string | null;
  sender_or_author: string | null;
  recipient_or_subject: string | null;
  location: string | null;
  people_mentioned: string[];
  places_mentioned: string[];
  organizations_mentioned: string[];
  full_transcription: string;
  subject_summary: string;
  keywords: string[];
  custom_fields: Record<string, unknown>;
  noteworthy_flag: boolean;
  noteworthy_reason: string | null;
  bibliography_entry: string;
  protocol_notes: string;
  confidence_score: string;
  raw_text?: string;
  needs_review?: boolean;
  processed_at?: string;
  filename?: string;
  status?: string;
  docx_path?: string;
}

export type ArchiveItem = MachensItem | GenericItem;

export interface AnalysisResult {
  id: string;
  project_id: string;
  group_label: string;
  criteria: string;
  item_count: number;
  created_at: string;
  docx_path?: string;
  data: Record<string, unknown>;
}

export interface Project {
  id: string;
  name: string;
  type: string;
  description: string;
  primary_language: string;
  date_range_from: string;
  date_range_to: string;
  primary_researcher: string;
  institution: string;
  notes: string;
  created_at: string;
  last_active: string;
}

export interface ProtocolVersion {
  version: number;
  content: string;
  saved_at: string;
}

export interface SchemaVersion {
  version: number;
  item_schema: string;
  analysis_schema: string;
  saved_at: string;
}

interface PatternEntry {
  count: number;
  projects: Set<string>;
  items: string[];
  dates?: string[];
  categories?: string[];
}

interface IntelligenceBus {
  projects: Record<string, Project>;
  activeProject: string | null;
  itemStore: ArchiveItem[];
  analysisStore: AnalysisResult[];
  patternLibrary: {
    people: Record<string, PatternEntry>;
    places: Record<string, PatternEntry>;
    themes: Record<string, PatternEntry>;
    dates: Record<string, PatternEntry>;
    categories: Record<string, PatternEntry>;
  };
  crossProjectIndex: {
    sharedPeople: Record<string, string[]>;
    sharedPlaces: Record<string, string[]>;
    dateOverlaps: Record<string, string[]>;
    thematicOverlaps: Record<string, string[]>;
  };
  sessionLog: Array<{ timestamp: string; action: string; item_id?: string; project_id?: string; details?: string }>;
  schemaRegistry: Record<string, SchemaVersion[]>;
  protocolRegistry: Record<string, ProtocolVersion[]>;
  exportQueue: string[];
}

const MACHEN_PROJECT_ID = "machen-family-papers";

const MACHEN_ITEM_SCHEMA = JSON.stringify({
  item_id: "UUID (auto-assigned)",
  project_id: "string",
  source_year: "string",
  source_type: "diary_typescript | diary_ocr_scan | holograph_letter | typescript_letter | news_clipping | photograph | postcard | ephemera | other",
  pass_applied: "pass_1 | pass_2 | pass_3 | pass_1_and_2 | all_three",
  register_rows: [
    {
      date: "string",
      person: "string (multiple: ' / ' separator)",
      category: "Vital | Residence | Education | Employment | Religion | Health | Travel | Relationship | Property | Cultural | Civic | Social | Voice | Provenance",
      fact: "string (self-contained; direct quote for Voice rows)",
      source: "string (citation with donee-anchor where applicable)"
    }
  ],
  donee_priority_categories: ["A", "B", "C", "D", "E"],
  headline_source_flag: "boolean",
  pass_1_coverage_report: "string or null",
  pass_2_anomalies: {
    category_1_ocr_errors: [],
    category_2_multiword_confusion: [],
    category_3_name_ambiguities: [],
    category_4_lacunae: [],
    category_6_period_usages_preserved: [],
    category_7_substantive_flags: []
  },
  noteworthy_flag: "boolean",
  noteworthy_reason: "string or null",
  strengthens_case: [],
  people_extracted: [],
  places_extracted: [],
  dates_extracted: [],
  confidence_score: "high | medium | low",
  ocr_correction_notes: "string"
}, null, 2);

const MACHEN_ANALYSIS_SCHEMA = JSON.stringify({
  group_label: "string",
  year_file_name: "Biographical_Facts_Register_Diary_[year].docx",
  item_count: "number",
  register_row_count: "number",
  narrative_diary: "string — flowing chronological prose",
  key_events: [{ date: "", event: "", donee_relevance: "" }],
  persons_directory: [{ name: "", role: "", priority_category: "", mentions: "" }],
  places_directory: [{ place: "", context: "", mentions: "" }],
  donee_summary: {
    LOC_mss86777_rows: 0,
    Westminster_rows: 0,
    Princeton_Wilson_rows: 0,
    Hopkins_rows: 0,
    other_rows: 0
  },
  voice_rows: [],
  ocr_corrections_applied: [],
  period_usages_preserved: [],
  shorthand_lacunae: [],
  bibliography: [],
  archivist_summary: "string",
  gaps_and_unknowns: "string",
  cross_reference_notes: "string"
}, null, 2);

const GENERIC_ITEM_SCHEMA = JSON.stringify({
  item_id: "UUID (auto-assigned)",
  project_id: "string",
  item_type: "string",
  date_on_item: "string or null",
  sender_or_author: "string or null",
  recipient_or_subject: "string or null",
  location: "string or null",
  people_mentioned: [],
  places_mentioned: [],
  organizations_mentioned: [],
  full_transcription: "string",
  subject_summary: "string",
  keywords: [],
  custom_fields: {},
  noteworthy_flag: "boolean",
  noteworthy_reason: "string or null",
  bibliography_entry: "string",
  protocol_notes: "string",
  confidence_score: "high | medium | low"
}, null, 2);

const GENERIC_ANALYSIS_SCHEMA = JSON.stringify({
  group_label: "string",
  group_criteria: "string",
  item_count: "number",
  date_range: "string",
  narrative_summary: "string",
  key_events: [],
  persons_directory: [],
  places_directory: [],
  themes: [],
  noteworthy_items: [],
  bibliography: [],
  archivist_summary: "string",
  gaps_and_unknowns: "string"
}, null, 2);

export const MACHEN_PROTOCOL = `MACHEN FAMILY PAPERS — DIARY PROCESSING PROTOCOL V.3
Compiled by Lyndon W. Cook, Hendersonville, Tennessee, May 2026.

SOURCE DESIGNATION: The Minnie Gresham Machen daily diary is a HEADLINE SOURCE. Voice-category extractions preserve direct quotation of Minnie's actual prose with sufficient context to be comprehensible without the source.

SOURCE HIERARCHY:
- Level 2 (Canonical): Typescript transcription. THE CANONICAL SOURCE.
- Level 3 (Working): OCR-scanned digital version. One degree removed. Subject to OCR errors the typescript does not contain.
- Master Transcript: Project owner's clean corrected version.
- No holograph diary volumes extant. Verification = OCR-correction-against-master-transcript, NOT holograph-verification.

THREE-PASS WORKFLOW:

PASS 1 — COVERAGE VERIFICATION:
Day-by-day verification the typescript covers the calendar year. Track across upload batches. Identify: missing days, duplicate-with-mis-dating sections, missing date headers, brief entries. Day-of-week verification is the standard tool for date-header ambiguity. Produce coverage report by month with cumulative-coverage statement.

PASS 2 — OCR / ANOMALIES CATALOG:
Category 1 — Probable OCR errors (rn/m, cl/d, ff/pf confusion, broken words at line breaks). Known corrections: Posher's→Posner's; Beney. Soc.→Benev. Soc.; Hambelton→Hambleton; Mt. Roland→Mt. Royal; Brook Bird→Brooke Bird; Von Kappf→Von Kapff
Category 2 — Multi-word confusion. Roland/Rowland/Royal triple confusion is a documented recurring Baltimore-area pattern — flag all three when any appears. Verify all three.
Category 3 — Name and term ambiguities. Flag for verification.
Category 4 — Substantive lacunae. Shorthand passages → preserve as Provenance rows. Pattern: Minnie used shorthand for private content.
Category 5 — Coverage gaps (cross-reference Pass 1)
Category 6 — Period usages preserved as authentic (do NOT modernize): "Telegram" (Rowland Multiplex device), "type-written" (period hyphenation), "Indian Ter." (1902 abbrev), "Aunty" (family naming), "wheel-ride" (period bicycle usage)
Category 7 — Substantive content flags for Register attention

STANDARDIZATION vs PRESERVATION:
STANDARDIZE: OCR introduced corruption the typescript itself does not contain.
PRESERVE: Typescript reads the word as Minnie actually wrote it.
Ambiguous cases → flagged for project-owner resolution.

PASS 3 — REGISTER EXTRACTION:
Build Register row blocks at day-and-event level. One entry = 0–8 rows depending on content density. Apply donee-locator priority to shape extraction granularity. Voice rows = direct quotation, never paraphrase.

DONEE-LOCATOR PRIORITY:
Category A (Highest — LOC mss86777): Minnie's own voice, Loy Gresham references (every mention), Generation II Gresham-side family network, Arthur Sr. daily life and legal practice
Category B (High — Westminster): J. Gresham Machen biographical material (every activity, decision, illness, social engagement), ecclesiastical and theological trajectory, 1925–1928 Princeton Seminary controversies
Category C (High — Princeton/Wilson): Wilson dinner attendance, October 1902 Princeton presidential inauguration, Wilson and Patton references, 1894 Wilson holograph acceptance letter context
Category D (Standard): General biographical material
Category E (Specialized): Hopkins University Special Collections; Mrs. Rowland connection; Confederate Bazaar Georgia Table; Phinizy/Augusta GA connection; Mrs. Burton late-life 1925–1928

DONEE SEQUENCING PRIORITY:
Tier 1: 1899-1901; 1928-1925 (reverse); 1913
Tier 2: 1894-1898; 1914-1916; 1917-1924
Tier 3: 1890-1893; 1903-1912
Prototype: 1902

CITATION CONVENTIONS:
Standard: "Minnie diary [year], [date] entry"
Multi-day: "Minnie diary [year], [date] entry ([event anchor])"
Donee-anchored: "[citation] (strengthens [Westminster/Princeton-Wilson/Hopkins/LOC] case)"
First-order: "Anderson, The Sharples-Sharpless Family, entry 9024"
Second-order: "[source] (via Suppl. A § 1)" or "(via Master Doc v.2)"
Project synthesis: cite "Project synthesis."

VERSIONING BY ADDITION: Entries never deleted — superseded by clarifying entries with explicit cross-reference.

PHASE A SKIP: Do NOT do Master-Document-narrative back-extraction before the typescript. Work is deferred to a single proper pass at first-order primary source.`;

export const BUS: IntelligenceBus = {
  projects: {},
  activeProject: null,
  itemStore: [],
  analysisStore: [],
  patternLibrary: {
    people: {},
    places: {},
    themes: {},
    dates: {},
    categories: {},
  },
  crossProjectIndex: {
    sharedPeople: {},
    sharedPlaces: {},
    dateOverlaps: {},
    thematicOverlaps: {},
  },
  sessionLog: [],
  schemaRegistry: {},
  protocolRegistry: {},
  exportQueue: [],
};

function initMachenProject(): void {
  const p: Project = {
    id: MACHEN_PROJECT_ID,
    name: "Machen Family Papers — Cook Collection",
    type: "Personal Diary / Family Papers",
    description: "Minnie Gresham Machen daily diary corpus 1890–1928. Cook collection. Compiled by Lyndon W. Cook, Hendersonville TN.",
    primary_language: "English",
    date_range_from: "1890",
    date_range_to: "1931",
    primary_researcher: "Lyndon W. Cook",
    institution: "Westminster Theological Seminary / Library of Congress / Princeton University Library",
    notes: "",
    created_at: new Date().toISOString(),
    last_active: new Date().toISOString(),
  };
  BUS.projects[MACHEN_PROJECT_ID] = p;
  BUS.activeProject = MACHEN_PROJECT_ID;

  BUS.protocolRegistry[MACHEN_PROJECT_ID] = [{
    version: 1,
    content: MACHEN_PROTOCOL,
    saved_at: new Date().toISOString(),
  }];

  BUS.schemaRegistry[MACHEN_PROJECT_ID] = [{
    version: 1,
    item_schema: MACHEN_ITEM_SCHEMA,
    analysis_schema: MACHEN_ANALYSIS_SCHEMA,
    saved_at: new Date().toISOString(),
  }];
}

initMachenProject();

export function isMachenProject(projectId: string): boolean {
  return projectId === MACHEN_PROJECT_ID || (BUS.projects[projectId]?.type ?? "").includes("Diary");
}

export function getActiveProject(): Project | null {
  if (!BUS.activeProject) return null;
  return BUS.projects[BUS.activeProject] ?? null;
}

export function getProjectItems(projectId: string): ArchiveItem[] {
  return BUS.itemStore.filter(i => i.project_id === projectId);
}

export function getProjectAnalysis(projectId: string): AnalysisResult[] {
  return BUS.analysisStore.filter(a => a.project_id === projectId);
}

export function getDefaultItemSchema(projectType: string): string {
  if (projectType.includes("Diary") || projectType.includes("Family")) {
    return MACHEN_ITEM_SCHEMA;
  }
  return GENERIC_ITEM_SCHEMA;
}

export function getDefaultAnalysisSchema(projectType: string): string {
  if (projectType.includes("Diary") || projectType.includes("Family")) {
    return MACHEN_ANALYSIS_SCHEMA;
  }
  return GENERIC_ANALYSIS_SCHEMA;
}

export function updatePatternLibrary(item: ArchiveItem): void {
  const projectId = item.project_id;
  const itemId = item.item_id;

  const people = (item as MachensItem).people_extracted ?? (item as GenericItem).people_mentioned ?? [];
  const places = (item as MachensItem).places_extracted ?? (item as GenericItem).places_mentioned ?? [];
  const dates = (item as MachensItem).dates_extracted ?? [];
  const keywords = (item as GenericItem).keywords ?? [];
  const rows = (item as MachensItem).register_rows ?? [];

  for (const p of people) {
    const key = p.toLowerCase().trim();
    if (!key) continue;
    if (!BUS.patternLibrary.people[key]) {
      BUS.patternLibrary.people[key] = { count: 0, projects: new Set(), items: [] };
    }
    BUS.patternLibrary.people[key].count++;
    BUS.patternLibrary.people[key].projects.add(projectId);
    BUS.patternLibrary.people[key].items.push(itemId);
  }

  for (const pl of places) {
    const key = pl.toLowerCase().trim();
    if (!key) continue;
    if (!BUS.patternLibrary.places[key]) {
      BUS.patternLibrary.places[key] = { count: 0, projects: new Set(), items: [] };
    }
    BUS.patternLibrary.places[key].count++;
    BUS.patternLibrary.places[key].projects.add(projectId);
    BUS.patternLibrary.places[key].items.push(itemId);
  }

  for (const d of dates) {
    const year = d.substring(0, 4);
    if (!BUS.patternLibrary.dates[year]) {
      BUS.patternLibrary.dates[year] = { count: 0, projects: new Set(), items: [] };
    }
    BUS.patternLibrary.dates[year].count++;
    BUS.patternLibrary.dates[year].projects.add(projectId);
    BUS.patternLibrary.dates[year].items.push(itemId);
  }

  for (const k of keywords) {
    const key = k.toLowerCase().trim();
    if (!key) continue;
    if (!BUS.patternLibrary.themes[key]) {
      BUS.patternLibrary.themes[key] = { count: 0, projects: new Set(), items: [] };
    }
    BUS.patternLibrary.themes[key].count++;
    BUS.patternLibrary.themes[key].projects.add(projectId);
    BUS.patternLibrary.themes[key].items.push(itemId);
  }

  for (const row of rows) {
    const cat = row.category;
    if (!cat) continue;
    if (!BUS.patternLibrary.categories[cat]) {
      BUS.patternLibrary.categories[cat] = { count: 0, projects: new Set(), items: [] };
    }
    BUS.patternLibrary.categories[cat].count++;
    BUS.patternLibrary.categories[cat].projects.add(projectId);
    BUS.patternLibrary.categories[cat].items.push(itemId);
  }

  for (const [personKey, entry] of Object.entries(BUS.patternLibrary.people)) {
    if (entry.projects.size > 1) {
      BUS.crossProjectIndex.sharedPeople[personKey] = Array.from(entry.projects);
    }
  }
  for (const [placeKey, entry] of Object.entries(BUS.patternLibrary.places)) {
    if (entry.projects.size > 1) {
      BUS.crossProjectIndex.sharedPlaces[placeKey] = Array.from(entry.projects);
    }
  }

  BUS.sessionLog.push({
    timestamp: new Date().toISOString(),
    action: "item_processed",
    item_id: itemId,
    project_id: projectId,
  });
}

export function getBusSummary() {
  const activeProj = BUS.activeProject ? BUS.projects[BUS.activeProject] : null;
  const activeItems = BUS.activeProject ? getProjectItems(BUS.activeProject) : [];
  const totalRows = activeItems.reduce((sum, item) => {
    return sum + ((item as MachensItem).register_rows?.length ?? 0);
  }, 0);

  return {
    activeProject: activeProj?.name ?? "None",
    activeProjectId: BUS.activeProject,
    projectCount: Object.keys(BUS.projects).length,
    totalItems: BUS.itemStore.length,
    activeProjectItems: activeItems.length,
    totalRows,
    patternCount: Object.keys(BUS.patternLibrary.people).length +
      Object.keys(BUS.patternLibrary.places).length,
    apiKeyActive: !!process.env["ANTHROPIC_API_KEY"],
  };
}

export function createProject(data: Omit<Project, "id" | "created_at" | "last_active">): Project {
  const id = uuidv4();
  const project: Project = {
    ...data,
    id,
    created_at: new Date().toISOString(),
    last_active: new Date().toISOString(),
  };
  BUS.projects[id] = project;

  const itemSchema = getDefaultItemSchema(data.type);
  const analysisSchema = getDefaultAnalysisSchema(data.type);
  BUS.schemaRegistry[id] = [{
    version: 1,
    item_schema: itemSchema,
    analysis_schema: analysisSchema,
    saved_at: new Date().toISOString(),
  }];
  BUS.protocolRegistry[id] = [];

  BUS.sessionLog.push({
    timestamp: new Date().toISOString(),
    action: "project_created",
    project_id: id,
    details: data.name,
  });

  return project;
}
