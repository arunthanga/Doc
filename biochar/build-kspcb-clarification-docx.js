/*
 * Formal letter (Word .docx) — request to KSPCB for a written clarification on the
 * applicable pollution-category (Orange vs Green) for a proposed agri-biomass
 * biochar (soil-amendment) unit at Eruthenpathy, Chittur, Palakkad, under the
 * KSPCB revised categorisation list as on 29.12.2025.
 *
 * Built with docx-js per the Anthropic "docx" skill.
 * Run:  NODE_PATH=../petitions/node_modules node build-kspcb-clarification-docx.js
 * Out:  KSPCB-Category-Clarification-Biochar.docx
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Footer, AlignmentType, HeadingLevel, LevelFormat, WidthType,
  ShadingType, BorderStyle, PageNumber, TabStopType,
} = require("docx");

const INK = "1A1A1A";
const ACCENT = "0B3D0B";
const RULE = "808080";
const HEADSHADE = "E9EFE9";
const ORANGEC = "C87A00";
const GREENC = "1E7D1E";
const CONTENT_W = 9638; // A4 content width @ 1134 DXA margins

const t = (text, opts = {}) => new TextRun({ text, font: "Times New Roman", size: 22, color: INK, ...opts });
const b = (text, opts = {}) => t(text, { bold: true, ...opts });

function para(children, opts = {}) {
  return new Paragraph({ spacing: { after: 120, line: 276 }, children: Array.isArray(children) ? children : [children], ...opts });
}
function labelLine(label, value) {
  return new Paragraph({ spacing: { after: 60, line: 264 }, children: [b(label + "  "), ...(Array.isArray(value) ? value : [t(value)])] });
}
function h(num, text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 90 },
    children: [new TextRun({ text: `${num}.  ${text}`, font: "Times New Roman", size: 23, bold: true, color: ACCENT })] });
}
function numItem(runs, ref = "asks") {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 100, line: 276 }, children: Array.isArray(runs) ? runs : [runs] });
}
function bullet(runs) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80, line: 268 }, children: Array.isArray(runs) ? runs : [runs] });
}
function thinRule() {
  return new Paragraph({ spacing: { before: 40, after: 120 }, border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 6, space: 1 } }, children: [t("")] });
}
function cell(children, { widthDxa, shade, header = false, align = AlignmentType.LEFT, size = 20 } = {}) {
  const kids = (Array.isArray(children) ? children : [children]).map((c) =>
    typeof c === "string"
      ? new Paragraph({ alignment: align, spacing: { after: 30, line: 250 },
          children: [new TextRun({ text: c, font: "Times New Roman", size, bold: header, color: header ? "222222" : INK })] })
      : c);
  return new TableCell({ width: { size: widthDxa, type: WidthType.DXA }, margins: { top: 50, bottom: 50, left: 90, right: 90 },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: "auto" } : undefined, children: kids });
}
function catRun(text, color) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 30, line: 250 },
    children: [new TextRun({ text, font: "Times New Roman", size: 20, bold: true, color })] });
}
function table(columnWidths, rows) {
  return new Table({ columnWidths, width: { size: columnWidths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: RULE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      left: { style: BorderStyle.SINGLE, size: 4, color: RULE }, right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE }, insideVertical: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    }, rows });
}

const children = [];

/* letterhead */
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 16 },
  children: [new TextRun({ text: "ARUN AND FRIENDS", font: "Times New Roman", size: 28, bold: true, color: INK })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 4 },
  children: [new TextRun({ text: "(Proposed partnership enterprise — Tender-Coconut & Agri-Biomass Biochar Unit)", font: "Times New Roman", size: 19, italics: true, color: "555555" })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 2 },
  children: [new TextRun({ text: "4/296, Mangapallam South, RVP Pudur Post, Palakkad District, Kerala \u2013 678555", font: "Times New Roman", size: 19, color: "555555" })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
  children: [new TextRun({ text: "Phone: 9901955667  \u00B7  Email: er.arunt@gmail.com", font: "Times New Roman", size: 19, color: "555555" })] }));
children.push(thinRule());

/* ref + date */
children.push(new Paragraph({ spacing: { after: 60 },
  tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
  children: [b("Ref: "), t("AAF/KSPCB/CAT/2026/[__]"), new TextRun({ text: "\t", font: "Times New Roman", size: 22 }), b("Date: "), t("[__] July 2026")] }));

/* addressee */
children.push(labelLine("To:", [
  t("The Member Secretary,"), new TextRun({ break: 1 }),
  t("Kerala State Pollution Control Board,"), new TextRun({ break: 1 }),
  t("Pattom, Thiruvananthapuram \u2013 695004."),
]));
children.push(labelLine("Copy to:", [
  t("The Environmental Engineer, KSPCB District Office, Palakkad."),
]));

/* subject */
children.push(new Paragraph({ spacing: { before: 60, after: 60, line: 276 }, children: [
  b("Subject:  "),
  t("Request for written clarification on the applicable pollution-control category (Orange vs. Green) for a proposed continuous, low-emission "),
  b("agri-biomass biochar (soil-amendment) unit"),
  t(" at Eruthenpathy, Chittur, Palakkad, under the KSPCB revised categorisation list as on 29.12.2025 \u2014 prior to filing the Consent to Establish."),
]}));

/* references */
children.push(new Paragraph({ spacing: { after: 40 }, children: [b("Reference:")] }));
children.push(bullet([t("KSPCB Notification No. KSPCB/159/2022-SEE-3 dated 10.08.2024 (Categorization Committee, 29.05.2024) \u2014 \u201CCharcoal Production from Coconut Shell by Pyrolysis \u2014 Orange.\u201D")]));
children.push(bullet([t("KSPCB Revised categorisation list as on 29.12.2025.")]));
children.push(bullet([t("KSPCB \u201CEnvironmental Guidelines for Charcoal Manufacturing Units\u201D dated 02.12.2025.")]));
children.push(bullet([t("KSPCB circular dated 16.09.2025 on self-certification-based CTE/CTO.")]));
children.push(bullet([t("CPCB Directions dated 07.03.2016 (u/s 18(1)(b) of the Water Act 1974 & Air Act 1981) and 12.02.2025 \u2014 Pollution-Index categorisation methodology.")]));

children.push(thinRule());

/* 1. respected sir */
children.push(new Paragraph({ spacing: { after: 100, line: 276 }, children: [t("Respected Sir/Madam,")] }));
children.push(para([
  t("We, "), b("Arun and Friends"),
  t(" (a proposed partnership enterprise), intend to establish a "),
  b("continuous, low-emission biochar unit"),
  t(" on an owned 38-cent plot at Eruthenpathy, Chittur Taluk (Block No. 30, Survey No. 300), Palakkad. Before filing our Consent to Establish, we respectfully seek a "),
  b("written clarification on the sector entry and pollution-category applicable to our unit"),
  t(" under the revised categorisation list as on 29.12.2025."),
]));

/* 2. the unit */
children.push(h(2, "Nature of the proposed unit"));
children.push(bullet([b("Product: "), t("biochar as a soil amendment / feedstock for enriched (microbe-inoculated) organic bio-fertilizer \u2014 "), b("not"), t(" fuel charcoal, "), b("not"), t(" activated carbon, and with "), b("no"), t(" pyrolysis-oil or carbon-black recovery.")]));
children.push(bullet([b("Feedstock: "), t("mixed agri-biomass \u2014 tender-coconut husk briquettes, coconut fronds/petioles, shells and coir pith, and seasonal agri-residue. It is not a coconut-shell-only charcoal operation.")]));
children.push(bullet([b("Capacity: "), t("Phase-1 \u2248 1 t/day biochar (consent applied rated for the 2 t/day ultimate).")]));
children.push(bullet([b("Process & controls: "), t("closed, continuous/semi-continuous pyrolysis with a syngas afterburner/flare (self-fuelling), cyclone/dust collector, alkaline scrubber with mist eliminator and a stack meeting KSPCB norms; continuous temperature monitoring and emergency shutdown; "), b("Zero-Liquid-Discharge"), t(" (tar condensate and scrubber blowdown captured and managed). No open burning at any stage.")]));

/* 3. classification question with candidate table */
children.push(h(3, "The classification question"));
children.push(para([
  t("Our unit could be read against either of two entries. We request confirmation of which one applies:"),
]));
{
  const cw = [4838, 1600, 3200];
  const rows = [
    new TableRow({ tableHeader: true, children: [
      cell("Candidate sector entry", { widthDxa: cw[0], shade: HEADSHADE, header: true }),
      cell("Category", { widthDxa: cw[1], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Consent regime", { widthDxa: cw[2], shade: HEADSHADE, header: true }),
    ]}),
    new TableRow({ children: [
      cell("\u201CCharcoal production from coconut shell by pyrolysis\u201D (Notification dated 10.08.2024)", { widthDxa: cw[0] }),
      cell([catRun("ORANGE", ORANGEC)], { widthDxa: cw[1] }),
      cell("Consent to Establish + Consent to Operate; periodic monitoring", { widthDxa: cw[2] }),
    ]}),
    new TableRow({ children: [
      cell("General charcoal / char (biochar) production \u2014 as we understand it is classified under the revised list as on 29.12.2025", { widthDxa: cw[0] }),
      cell([catRun("GREEN", GREENC)], { widthDxa: cw[1] }),
      cell("Simplified consent regime", { widthDxa: cw[2] }),
    ]}),
  ];
  children.push(table(cw, rows));
}
children.push(new Paragraph({ spacing: { after: 100 }, children: [t("")] }));

/* 4. why we submit the general/green entry fits */
children.push(h(4, "Why we submit the general (lighter) entry is the correct fit"));
children.push(bullet([b("Not \u201Ccoconut shell.\u201D "), t("Our feedstock is mixed agri-biomass, not a coconut-shell-only charge; the 10.08.2024 Orange entry is worded specifically for \u201Ccoconut shell.\u201D")]));
children.push(bullet([b("Not \u201Ccharcoal\u201D for fuel, nor activated carbon. "), t("The product is biochar for soil; there is no oil/carbon-black recovery and no chemical/steam activation \u2014 the features that drive the heavier categories.")]));
children.push(bullet([b("Low pollution potential on the CPCB Pollution-Index basis. "), t("A closed, continuous process with afterburner, scrubber, controlled stack emissions and Zero-Liquid-Discharge has a materially lower Pollution Index than the coconut-shell charcoal / activated-carbon activities.")]));
children.push(bullet([b("Guideline-compliant by design. "), t("The unit is designed to meet the \u201CEnvironmental Guidelines for Charcoal Manufacturing Units\u201D dated 02.12.2025 (emission control, stack, ash/dust housekeeping, heat reutilisation, responsible feedstock sourcing).")]));

/* 5. clarifications sought */
children.push(h(5, "Clarifications requested"));
children.push(para([t("We respectfully request the Board to clarify, in writing:")]));
children.push(numItem([b("The applicable entry and category. "), t("Whether the described agri-biomass biochar (soil-amendment) unit falls under the general charcoal/char (biochar) production entry (Green) in the revised list as on 29.12.2025, rather than the \u201Ccharcoal production from coconut shell by pyrolysis\u201D (Orange) entry of 10.08.2024.")]));
children.push(numItem([b("If Orange is held to apply, "), t("the consent conditions, and whether the unit is eligible for the self-certification route under the 16.09.2025 circular at our capital-investment level.")]));
children.push(numItem([b("The applicable technical norms "), t("(emission limits, minimum stack height, distance/siting criteria) under the 02.12.2025 charcoal guidelines that we must design to.")]));
children.push(numItem([b("Confirmation that no public hearing or Environmental Clearance "), t("is required for the unit (being outside the Red / EIA-2006 scheduled category).")]));

/* 6. enclosures */
children.push(h(6, "Enclosures"));
children.push(bullet([t("Brief project note / DPR extract (concept, capacity, mass balance).")]));
children.push(bullet([t("Process-flow and emission-control schematic (afterburner \u2192 cyclone \u2192 scrubber \u2192 stack; ZLD line).")]));
children.push(bullet([t("Feedstock list and plot details (survey no., extent, land-use, setbacks).")]));

/* closing */
children.push(new Paragraph({ spacing: { before: 160, after: 120, line: 276 }, children: [
  t("An early written clarification will let us file a correctly-categorised Consent to Establish and design the unit to the exact norms. We shall be glad to furnish any further particulars or meet the concerned officer at the District Office."),
]}));
children.push(new Paragraph({ spacing: { after: 60 }, children: [t("Thanking you,")] }));
children.push(new Paragraph({ spacing: { after: 200 }, children: [t("Yours faithfully,")] }));
children.push(new Paragraph({ spacing: { after: 16 }, children: [t("_________________________________")] }));
children.push(new Paragraph({ spacing: { after: 16 }, children: [b("Arun T."), t("  \u2014 for and on behalf of Arun and Friends (proposed)")] }));
children.push(new Paragraph({ spacing: { after: 60 }, children: [t("Phone: 9901955667  \u00B7  Email: er.arunt@gmail.com")] }));

const doc = new Document({
  creator: "Arun and Friends (proposed)",
  title: "KSPCB category clarification request — agri-biomass biochar unit, Chittur",
  description: "Written request to KSPCB for clarification of pollution-category (Orange vs Green) for a biomass biochar-for-soil unit.",
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 22, color: INK } } },
    paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: "Times New Roman", size: 23, bold: true, color: ACCENT }, paragraph: { spacing: { before: 200, after: 90 }, outlineLevel: 0 } }],
  },
  numbering: { config: [
    { reference: "asks", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 620, hanging: 340 } } } }] },
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 520, hanging: 260 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      children: [
        new TextRun({ text: "Arun and Friends \u2014 KSPCB category clarification (biochar unit, Chittur)", font: "Times New Roman", size: 16, color: "999999" }),
        new TextRun({ text: "\tPage ", font: "Times New Roman", size: 16, color: "999999" }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 16, color: "999999" }),
        new TextRun({ text: " of ", font: "Times New Roman", size: 16, color: "999999" }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Times New Roman", size: 16, color: "999999" }),
      ],
    })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = "KSPCB-Category-Clarification-Biochar.docx";
  fs.writeFileSync(out, buf);
  console.log("Wrote", out, buf.length, "bytes");
});
