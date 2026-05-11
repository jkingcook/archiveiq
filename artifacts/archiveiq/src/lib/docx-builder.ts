import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
  Footer, PageNumber, NumberFormat, Header,
} from "docx";
import { writeFileSync } from "fs";
import type { MachensItem, ArchiveItem, AnalysisResult } from "./intelligence-bus.js";

const GOLD = "C9A84C";
const DARK_BROWN = "3D2B1F";
const LIGHT_TAN = "FAF7F0";

function h1(text: string): Paragraph {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 120 },
  });
}

function h2(text: string): Paragraph {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
  });
}

function h3(text: string): Paragraph {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 60 },
  });
}

function para(text: string, opts?: { italic?: boolean; bold?: boolean; indent?: boolean }): Paragraph {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        italics: opts?.italic,
        bold: opts?.bold,
        size: 22,
      }),
    ],
    indent: opts?.indent ? { left: 720 } : undefined,
    spacing: { after: 80 },
  });
}

function divider(): Paragraph {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD } },
    spacing: { after: 120 },
    text: "",
  });
}

function buildRegisterTable(rows: MachensItem["register_rows"]): Table {
  const headerRow = new TableRow({
    tableHeader: true,
    children: ["Date", "Person", "Category", "Fact", "Source"].map(h =>
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 20 })],
        })],
        shading: { type: ShadingType.SOLID, color: DARK_BROWN },
        width: { size: 20, type: WidthType.PERCENTAGE },
      })
    ),
  });

  const dataRows = rows.map(row => {
    const isVoice = row.category === "Voice";
    const isHighDonee = row.source?.includes("strengthens");
    return new TableRow({
      children: [row.date, row.person, row.category, row.fact, row.source].map((val, i) =>
        new TableCell({
          children: [new Paragraph({
            children: [new TextRun({
              text: val ?? "",
              italics: isVoice && i === 3,
              size: 20,
            })],
            indent: isVoice && i === 3 ? { left: 360 } : undefined,
          })],
          shading: isHighDonee ? { type: ShadingType.SOLID, color: "FFF9E6" } : undefined,
          width: { size: 20, type: WidthType.PERCENTAGE },
        })
      ),
    });
  });

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

function buildGenericMetaTable(item: ArchiveItem): Table {
  type GenItem = { item_type?: string; date_on_item?: string; sender_or_author?: string; recipient_or_subject?: string; location?: string; confidence_score?: string; filename?: string };
  const g = item as GenItem;
  const rows: [string, string][] = [
    ["Item ID", item.item_id],
    ["Project", item.project_id],
    ["Type", g.item_type ?? ""],
    ["Date", g.date_on_item ?? ""],
    ["Sender / Author", g.sender_or_author ?? ""],
    ["Recipient / Subject", g.recipient_or_subject ?? ""],
    ["Location", g.location ?? ""],
    ["Confidence", g.confidence_score ?? ""],
    ["File", g.filename ?? ""],
    ["Processed", item.processed_at ?? ""],
  ];

  return new Table({
    rows: rows.map(([k, v]) => new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: k, bold: true, size: 20 })] })],
          shading: { type: ShadingType.SOLID, color: "F5EFE0" },
          width: { size: 30, type: WidthType.PERCENTAGE },
        }),
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: v, size: 20 })] })],
          width: { size: 70, type: WidthType.PERCENTAGE },
        }),
      ],
    })),
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

export async function buildMachenItemDocx(item: MachensItem, outPath: string): Promise<void> {
  const rows = item.register_rows ?? [];
  const anomalies = item.pass_2_anomalies ?? {};

  const children: Paragraph[] = [
    h1("Machen Family Papers — Biographical Facts Register"),
    divider(),
    h2("1. Coverage & Source Metadata"),
    para(`Item ID: ${item.item_id}`),
    para(`Source Year: ${item.source_year}`),
    para(`Source Type: ${item.source_type}`),
    para(`Pass Applied: ${item.pass_applied}`),
    para(`Headline Source: ${item.headline_source_flag ? "YES" : "NO"}`),
    para(`Donee Categories: ${(item.donee_priority_categories ?? []).join(", ")}`),
    para(`Confidence: ${item.confidence_score}`),
    para(`File: ${item.filename ?? ""}`),
    para(`Processed: ${item.processed_at ?? ""}`),
    divider(),
    h2("2. Pass 1 — Coverage Report"),
    para(item.pass_1_coverage_report ?? "Not applied."),
    divider(),
    h2("3. Pass 2 — Anomalies Catalog"),
    h3("Category 1 — Probable OCR Errors"),
    ...(anomalies.category_1_ocr_errors ?? []).map(s => para(`• ${s}`)),
    h3("Category 2 — Multi-Word Confusion"),
    ...(anomalies.category_2_multiword_confusion ?? []).map(s => para(`• ${s}`)),
    h3("Category 3 — Name Ambiguities"),
    ...(anomalies.category_3_name_ambiguities ?? []).map(s => para(`• ${s}`)),
    h3("Category 4 — Substantive Lacunae"),
    ...(anomalies.category_4_lacunae ?? []).map(s => para(`• ${s}`)),
    h3("Category 6 — Period Usages Preserved"),
    ...(anomalies.category_6_period_usages_preserved ?? []).map(s => para(`• ${s}`)),
    h3("Category 7 — Substantive Flags"),
    ...(anomalies.category_7_substantive_flags ?? []).map(s => para(`• ${s}`)),
    divider(),
    h2("4. Register Rows"),
  ];

  const doc = new Document({
    sections: [{
      properties: {},
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [new TextRun({ text: "Machen Family Papers — Biographical Facts Register", bold: true, color: DARK_BROWN, size: 18 })],
            alignment: AlignmentType.CENTER,
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "Machen Family Papers | Schema v.4 | Protocol v.3 | Processed: " + (item.processed_at ?? "") + " | Lyndon W. Cook", size: 16 }),
              new TextRun({ children: [" | Page ", new PageNumber()], size: 16 }),
            ],
          })],
        }),
      },
      children: [
        ...children,
        rows.length > 0 ? buildRegisterTable(rows) : para("No register rows extracted."),
        new Paragraph({ text: "", spacing: { after: 200 } }),
        divider(),
        h2("5. Donee Summary"),
        para(`People Extracted: ${(item.people_extracted ?? []).join("; ")}`),
        para(`Places Extracted: ${(item.places_extracted ?? []).join("; ")}`),
        para(`Dates Extracted: ${(item.dates_extracted ?? []).join("; ")}`),
        para(`Strengthens Case: ${(item.strengthens_case ?? []).join("; ")}`),
        divider(),
        h2("6. Noteworthy"),
        para(item.noteworthy_flag ? `YES — ${item.noteworthy_reason ?? ""}` : "Not flagged."),
        divider(),
        h2("7. OCR Correction Notes"),
        para(item.ocr_correction_notes ?? "None."),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  writeFileSync(outPath, buffer);
}

export async function buildGenericItemDocx(item: ArchiveItem, outPath: string): Promise<void> {
  type GenItem = { full_transcription?: string; subject_summary?: string; keywords?: string[]; bibliography_entry?: string; protocol_notes?: string; noteworthy_flag?: boolean; noteworthy_reason?: string; people_mentioned?: string[]; places_mentioned?: string[]; organizations_mentioned?: string[]; confidence_score?: string };
  const g = item as GenItem;

  const doc = new Document({
    sections: [{
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [new TextRun({ text: `ArchiveIQ | ${item.project_id} | ${item.processed_at ?? ""}`, size: 16 })],
          })],
        }),
      },
      children: [
        h1("ArchiveIQ — Archival Item Record"),
        divider(),
        h2("1. Item Metadata"),
        buildGenericMetaTable(item),
        new Paragraph({ text: "", spacing: { after: 200 } }),
        divider(),
        h2("2. Full Transcription"),
        para(g.full_transcription ?? ""),
        divider(),
        h2("3. People / Places / Organizations"),
        h3("People Mentioned"),
        ...(g.people_mentioned ?? []).map(s => para(`• ${s}`)),
        h3("Places Mentioned"),
        ...(g.places_mentioned ?? []).map(s => para(`• ${s}`)),
        h3("Organizations Mentioned"),
        ...(g.organizations_mentioned ?? []).map(s => para(`• ${s}`)),
        divider(),
        h2("4. Subject Summary"),
        para(g.subject_summary ?? ""),
        divider(),
        h2("5. Keywords & Themes"),
        para((g.keywords ?? []).join(", ")),
        divider(),
        h2("6. Bibliography Entry"),
        para(g.bibliography_entry ?? "", { italic: true }),
        divider(),
        h2("7. Protocol Notes"),
        para(g.protocol_notes ?? ""),
        divider(),
        h2("8. Noteworthy"),
        para(g.noteworthy_flag ? `YES — ${g.noteworthy_reason ?? ""}` : "Not flagged."),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  writeFileSync(outPath, buffer);
}

export async function buildGroupAnalysisDocx(
  analysis: AnalysisResult,
  project: { name: string; type: string },
  outPath: string
): Promise<void> {
  const data = analysis.data as Record<string, unknown>;
  const isMachen = project.type.includes("Diary") || project.type.includes("Family");

  const children: Paragraph[] = [
    h1(isMachen ? "Biographical Facts Register" : "Collection Analysis Report"),
    para(analysis.group_label, { bold: true }),
    para(`${analysis.item_count} items | ${analysis.created_at}`),
    isMachen ? para("Machen Family Papers | Schema v.4 | Protocol v.3") : para(project.name),
    divider(),
    h2("1. Archivist Summary"),
    para(String(data.archivist_summary ?? data.narrative_summary ?? "")),
    divider(),
    h2("2. Narrative"),
    para(String(data.narrative_diary ?? data.narrative_summary ?? "")),
    divider(),
    h2("3. Key Events"),
    ...((data.key_events as Array<{ date: string; event: string; donee_relevance?: string }>) ?? []).map(e =>
      para(`${e.date}: ${e.event}${e.donee_relevance ? " [" + e.donee_relevance + "]" : ""}`)
    ),
    divider(),
    h2("4. Persons Directory"),
    ...((data.persons_directory as Array<{ name: string; role: string; priority_category?: string; mentions?: string }>) ?? []).map(p =>
      para(`${p.name} — ${p.role}${p.priority_category ? " [Cat. " + p.priority_category + "]" : ""}${p.mentions ? " (" + p.mentions + " mentions)" : ""}`)
    ),
    divider(),
    h2("5. Places Directory"),
    ...((data.places_directory as Array<{ place: string; context: string }>) ?? []).map(p =>
      para(`${p.place}: ${p.context}`)
    ),
    divider(),
    h2("6. Gaps & Unknowns"),
    para(String(data.gaps_and_unknowns ?? "")),
  ];

  if (isMachen) {
    const ds = data.donee_summary as Record<string, number> | undefined;
    if (ds) {
      children.push(divider(), h2("7. Donee Summary"));
      for (const [k, v] of Object.entries(ds)) {
        children.push(para(`${k}: ${v} rows`));
      }
    }
    children.push(divider(), h2("8. Bibliography"));
    for (const b of (data.bibliography as string[]) ?? []) {
      children.push(para(b, { italic: true }));
    }
    children.push(divider(), h2("9. Cross-Reference Notes"));
    children.push(para(String(data.cross_reference_notes ?? "")));
  } else {
    children.push(divider(), h2("7. Themes"));
    for (const t of (data.themes as string[]) ?? []) {
      children.push(para(`• ${t}`));
    }
    children.push(divider(), h2("8. Bibliography"));
    for (const b of (data.bibliography as string[]) ?? []) {
      children.push(para(b, { italic: true }));
    }
  }

  const doc = new Document({
    sections: [{
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [new TextRun({
              text: isMachen
                ? `Biographical Facts Register | Machen Family Papers | Schema v.4 + Protocol v.3 | Lyndon W. Cook`
                : `ArchiveIQ | ${project.name} | ${analysis.created_at}`,
              size: 16,
            })],
          })],
        }),
      },
      children,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  writeFileSync(outPath, buffer);
}

export async function buildMasterReportDocx(
  projectId: string,
  project: { name: string; type: string; primary_researcher?: string; description?: string },
  items: ArchiveItem[],
  analyses: AnalysisResult[],
  outPath: string
): Promise<void> {
  const isMachen = project.type.includes("Diary") || project.type.includes("Family");
  const now = new Date().toISOString().split("T")[0];

  const allPeople = new Set<string>();
  const allPlaces = new Set<string>();
  let totalRows = 0;

  for (const item of items) {
    const m = item as MachensItem;
    for (const p of m.people_extracted ?? (item as { people_mentioned?: string[] }).people_mentioned ?? []) allPeople.add(p);
    for (const p of m.places_extracted ?? (item as { places_mentioned?: string[] }).places_mentioned ?? []) allPlaces.add(p);
    totalRows += m.register_rows?.length ?? 0;
  }

  const noteworthy = items.filter(i => (i as MachensItem).noteworthy_flag);

  const children: Paragraph[] = [
    new Paragraph({
      children: [new TextRun({ text: isMachen ? "The Cook Collection — Machen Family Papers" : project.name, bold: true, size: 36, color: DARK_BROWN })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: isMachen ? "Complete Biographical Facts Register" : "Master Collection Report", size: 28 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `Total Items: ${items.length} | Total Rows: ${totalRows}`, size: 22 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: isMachen ? `Compiled by ${project.primary_researcher ?? "Lyndon W. Cook"}, Hendersonville, Tennessee` : (project.primary_researcher ?? ""), size: 22 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: isMachen ? `Schema v.4 + Diary Processing Protocol v.3 | May 2026` : now, size: 22 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    divider(),
    h2("1. Collection Executive Summary"),
    para(project.description ?? ""),
    para(`${items.length} items processed across ${analyses.length} analysis group(s).`),
    para(`Total Register Rows: ${totalRows} | People Identified: ${allPeople.size} | Places Identified: ${allPlaces.size}`),
    divider(),
    h2("2. Collection Statistics"),
    para(`Items processed: ${items.length}`),
    para(`Analysis groups: ${analyses.length}`),
    para(`Unique people: ${allPeople.size}`),
    para(`Unique places: ${allPlaces.size}`),
    para(`Noteworthy items: ${noteworthy.length}`),
    divider(),
    h2("3. Complete Persons Index (A–Z)"),
    ...[...allPeople].sort().map(p => para(`• ${p}`)),
    divider(),
    h2("4. Complete Places Index"),
    ...[...allPlaces].sort().map(p => para(`• ${p}`)),
    divider(),
    h2("5. All Noteworthy Items"),
    ...noteworthy.map(i => para(`${i.item_id}: ${(i as MachensItem).noteworthy_reason ?? ""}`)),
  ];

  if (isMachen) {
    const allVoice: MachensItem["register_rows"] = [];
    for (const item of items) {
      const m = item as MachensItem;
      for (const row of m.register_rows ?? []) {
        if (row.category === "Voice") allVoice.push(row);
      }
    }
    if (allVoice.length > 0) {
      children.push(divider(), h2("6. All Voice Rows (Headline-Source Quotations)"));
      children.push(buildRegisterTable(allVoice));
      children.push(new Paragraph({ text: "", spacing: { after: 200 } }));
    }
  }

  const doc = new Document({
    sections: [{ children }],
  });
  const buffer = await Packer.toBuffer(doc);
  writeFileSync(outPath, buffer);
}
