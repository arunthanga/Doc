/*
 * Formal petition (Word .docx) — Pollution-category parity for the coir industry in Chittur.
 * Built with docx-js per the Anthropic "docx" skill (explicit A4 size, dual-width tables,
 * numbering config for lists, CLEAR shading, page breaks inside paragraphs).
 *
 * Run:  node build-coir-petition-docx.js
 * Out:  Coir-Industry-Pollution-Category-Petition.docx
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat, WidthType,
  ShadingType, BorderStyle, PageNumber, TabStopType, TabStopPosition,
} = require("docx");

/* ---- palette / metrics ---------------------------------------------------- */
const INK = "1A1A1A";
const GREEN = "0B3D0B";        // section headings
const RULE = "808080";
const HEADSHADE = "E9EFE9";    // table header shade
const REDC = "C0392B";
const ORANGEC = "C87A00";
const GREENC = "1E7D1E";
const WHITEC = "444444";

// A4 content width = 11906 - (2 * 1134) margins = 9638 DXA
const CONTENT_W = 9638;

/* ---- small builders ------------------------------------------------------- */
const t = (text, opts = {}) => new TextRun({ text, font: "Times New Roman", size: 22, color: INK, ...opts });
const b = (text, opts = {}) => t(text, { bold: true, ...opts });

function para(children, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    children: Array.isArray(children) ? children : [children],
    ...opts,
  });
}

function labelLine(label, value) {
  return new Paragraph({
    spacing: { after: 60, line: 264 },
    children: [b(label + "  "), ...(Array.isArray(value) ? value : [t(value)])],
  });
}

function h(num, text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text: `${num}.  ${text}`, font: "Times New Roman", size: 24, bold: true, color: GREEN })],
  });
}

function annexHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 160, after: 120 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true, color: GREEN })],
  });
}

function numItem(runs, ref = "petition-list") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 100, line: 276 },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function bullet(runs) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80, line: 268 },
    children: Array.isArray(runs) ? runs : [runs],
  });
}

function pageBreakPara() {
  return new Paragraph({ children: [new TextRun({ break: 1 })], pageBreakBefore: true });
}

function thinRule() {
  return new Paragraph({
    spacing: { before: 40, after: 120 },
    border: { bottom: { color: RULE, style: BorderStyle.SINGLE, size: 6, space: 1 } },
    children: [t("")],
  });
}

/* ---- table helpers (dual widths, CLEAR shading) --------------------------- */
function cell(children, { widthDxa, shade, header = false, align = AlignmentType.LEFT } = {}) {
  const kids = (Array.isArray(children) ? children : [children]).map((c) =>
    typeof c === "string"
      ? new Paragraph({
          alignment: align,
          spacing: { after: 40, line: 252 },
          children: [new TextRun({ text: c, font: "Times New Roman", size: 20, bold: header, color: header ? "222222" : INK })],
        })
      : c
  );
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: "auto" } : undefined,
    children: kids,
  });
}

function catRun(text, color) {
  return new Paragraph({
    spacing: { after: 40, line: 252 },
    children: [new TextRun({ text, font: "Times New Roman", size: 20, bold: true, color })],
  });
}

function table(columnWidths, rows) {
  return new Table({
    columnWidths,
    width: { size: columnWidths.reduce((a, c) => a + c, 0), type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      left: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    },
    rows,
  });
}

/* ========================================================================== */
/* CONTENT                                                                     */
/* ========================================================================== */
const children = [];

/* --- letterhead ----------------------------------------------------------- */
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 20 },
  children: [new TextRun({ text: "THE CHITTUR AGRICULTURE MARKETING SOCIETY, KOZHINJAMPARA", font: "Times New Roman", size: 26, bold: true, color: INK })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 40 },
  children: [new TextRun({ text: "Chittur Taluk, Palakkad District, Kerala", font: "Times New Roman", size: 20, italics: true, color: "555555" })],
}));
children.push(thinRule());

/* --- title ---------------------------------------------------------------- */
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 60, after: 60 },
  children: [new TextRun({ text: "PETITION FOR POLLUTION-CATEGORY PARITY FOR THE COIR INDUSTRY IN CHITTUR", font: "Times New Roman", size: 26, bold: true, color: INK })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 180 },
  children: [new TextRun({ text: "Re-alignment of husk retting and coir de-fibring / ginning with the CPCB Pollution-Index methodology and with Tamil Nadu's White-category treatment of the same husk and process chain", font: "Times New Roman", size: 21, italics: true, color: "444444" })],
}));

/* --- addressing block ----------------------------------------------------- */
children.push(labelLine("To:", [
  t("Shri P. K. Kunhalikutty,"), new TextRun({ break: 1 }),
  t("Hon'ble Minister for Industries & Commerce, Government of Kerala,"), new TextRun({ break: 1 }),
  t("Thiruvananthapuram."),
]));
children.push(labelLine("Copy to:", [
  t("1. Shri Ramesh Chennithala, Hon'ble Minister in charge of Coir."), new TextRun({ break: 1 }),
  t("2. The Chairman, Kerala State Pollution Control Board (KSPCB)."),
]));
children.push(labelLine("From:", [
  t("The Chittur Agriculture Marketing Society, Kozhinjampara, Chittur Taluk, Palakkad District — represented by "),
  b("[Name, designation]"), t(" and delegation."),
]));
children.push(labelLine("Through:", [t("Adv. Sumesh Achuthan, MLA, Chittur Constituency.")]));
children.push(labelLine("Date:", [b("[__] July 2026"), t(", Thiruvananthapuram.")]));
children.push(new Paragraph({
  spacing: { before: 40, after: 160, line: 276 },
  children: [
    b("Subject:  "),
    t("Request to re-align the pollution-control classification of coconut-husk retting (presently Red) and coir de-fibring / ginning (presently Orange) in Kerala with the CPCB Pollution-Index norms and with Tamil Nadu's White-category treatment of the identical process chain, so that the primary coir industry can develop within Chittur instead of migrating across the district border to Pollachi (Tamil Nadu)."),
  ],
}));
children.push(thinRule());

/* --- 1. Respected Sir ------------------------------------------------------ */
children.push(h(1, "Respected Sir"));
children.push(para([
  t("Chittur Taluk adjoins the Pollachi coir belt of Tamil Nadu and enjoys abundant coconut and coir-husk availability. Yet the "),
  b("primary coir-processing industry has migrated almost entirely across the border to Pollachi"),
  t(" — barely 80 km away — chiefly because of a difference in "),
  b("pollution-control classification"),
  t(" between the two States for one and the same raw material and process chain. We respectfully pray for a level playing field."),
]));
children.push(para([
  t("We wish, at the outset, to state the position "),
  b("precisely and without overstatement"),
  t(". On an earlier occasion the classification gap was described more broadly than the record supports. Having now examined KSPCB's current classification list, we place before the Government the "),
  b("exact and narrower position"),
  t(": Kerala already treats dry coir manufacturing leniently (White), in line with the national CPCB norm. The genuine bottleneck for a Chittur unit is sharp and limited — "),
  b("husk retting is in the Red category and coir de-fibring / ginning is in the Orange category"),
  t("."),
]));

/* --- 2. Precise regulatory position ---------------------------------------- */
children.push(h(2, "The precise regulatory position in Kerala (KSPCB current list, December 2025)"));
children.push(para([
  t("As per the Kerala State Pollution Control Board's current classification of industrial sectors, the coir process chain is placed as follows (reproduced in full at "),
  b("Annexure A"), t("):"),
]));

/* KSPCB summary table */
{
  const cw = [3900, 1900, 3838]; // = 9638
  const rows = [
    new TableRow({ tableHeader: true, children: [
      cell("Coir process / activity", { widthDxa: cw[0], shade: HEADSHADE, header: true }),
      cell("KSPCB category", { widthDxa: cw[1], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Consent regime & implication", { widthDxa: cw[2], shade: HEADSHADE, header: true }),
    ]}),
    new TableRow({ children: [
      cell("Coconut-husk retting", { widthDxa: cw[0] }),
      cell([catRun("RED", REDC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell("Highest scrutiny — full Consent to Establish (CTE) and Consent to Operate (CTO); can attract public-hearing-tier appraisal.", { widthDxa: cw[2] }),
    ]}),
    new TableRow({ children: [
      cell("Coir de-fibring / ginning; coir stencilling; coir bleaching / dyeing", { widthDxa: cw[0] }),
      cell([catRun("ORANGE", ORANGEC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell("Full CTE and CTO required, with periodic monitoring.", { widthDxa: cw[2] }),
    ]}),
    new TableRow({ children: [
      cell("Coir fibre / pith processing generating effluent", { widthDxa: cw[0] }),
      cell([catRun("GREEN", GREENC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell("Simplified consent regime.", { widthDxa: cw[2] }),
    ]}),
    new TableRow({ children: [
      cell("Dry coir processing / coir products without effluent; coir yarn without dyeing; coir-pith final product; coir rope; curled coir", { widthDxa: cw[0] }),
      cell([catRun("WHITE", WHITEC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell("Practically non-polluting; no Consent to Operate — a simple intimation to KSPCB suffices.", { widthDxa: cw[2] }),
    ]}),
  ];
  children.push(table(cw, rows));
}
children.push(new Paragraph({ spacing: { after: 120 }, children: [t("")] }));
children.push(para([
  t("Two points follow. "),
  b("First"), t(", Kerala "),
  b("already"), t(" treats dry coir manufacturing — coir products, yarn without dyeing, coir pith as a final product, rope and curled coir — as "),
  b("White"), t(", matching the national CPCB White-category treatment. There is therefore "),
  b("no dispute"), t(" about dry value-addition. "),
  b("Second"), t(", the real and only bottleneck is at the "),
  b("primary end"), t(" of the chain: retting (Red) and de-fibring / ginning (Orange). Both require full Consent to Establish and Consent to Operate, and retting specifically can trigger public-hearing-tier scrutiny — which, for small MSME/cottage units, is prohibitive."),
]));

/* --- 3. Comparator: Tamil Nadu -------------------------------------------- */
children.push(h(3, "The comparator: Tamil Nadu treats the identical husk and process as White"));
children.push(para([
  t("The Tamil Nadu Pollution Control Board (TNPCB) classifies the entire "),
  b("\u201Ccoconut-husk retting / de-fibering / pith-processing industry\u201D"),
  t(" in the "), b("White"), t(" category. The recent history is directly on point (reproduced at "),
  b("Annexure B"), t("):"),
]));
children.push(bullet([
  b("10-11-2021 — "),
  t("By proceedings dated 10 November 2021, TNPCB reclassified the coir industry from White to Orange, purportedly on directions of the environmental tribunal and courts, thereby subjecting these units to consent and periodic monitoring."),
]));
children.push(bullet([
  b("12-10-2023 — "),
  t("After an enormous number of representations from coir exporters' and manufacturers' associations, TNPCB formally "),
  b("withdrew"), t(" that Orange reclassification, expressly "),
  new TextRun({ text: "\u201Cconsidering the coir industry is one of the vital agro-based cottage-type industries in the State contributing significantly to the creation of livelihood in major coconut-growing districts and to encourage the cottage-type MSME sector for the production of eco-friendly products\u201D", font: "Times New Roman", size: 22, italics: true, color: INK }),
  t(", restoring the coconut-husk retting / de-fibring / pith-processing industry to the "),
  b("White"), t(" category."),
]));
children.push(para([
  t("The comparison could not be starker. It is the "),
  b("same husk, the same process chain, roughly 80 km apart across the same district border"),
  t(" — yet placed in opposite regulatory tiers:"),
]));

/* KL vs TN comparison table */
{
  const cw = [3238, 3200, 3200]; // = 9638
  const rows = [
    new TableRow({ tableHeader: true, children: [
      cell("Process", { widthDxa: cw[0], shade: HEADSHADE, header: true }),
      cell("Kerala (KSPCB)", { widthDxa: cw[1], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Tamil Nadu (TNPCB)", { widthDxa: cw[2], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell("Coconut-husk retting", { widthDxa: cw[0] }),
      cell([catRun("RED", REDC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell([catRun("WHITE", WHITEC)], { widthDxa: cw[2], align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell("Coir de-fibring / ginning", { widthDxa: cw[0] }),
      cell([catRun("ORANGE", ORANGEC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell([catRun("WHITE", WHITEC)], { widthDxa: cw[2], align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      cell("Coir-pith processing", { widthDxa: cw[0] }),
      cell([catRun("GREEN (effluent) / WHITE (dry)", GREENC)], { widthDxa: cw[1], align: AlignmentType.CENTER }),
      cell([catRun("WHITE", WHITEC)], { widthDxa: cw[2], align: AlignmentType.CENTER }),
    ]}),
  ];
  children.push(table(cw, rows));
}
children.push(new Paragraph({ spacing: { after: 120 }, children: [t("")] }));

/* --- 4. Legal / policy basis ---------------------------------------------- */
children.push(h(4, "The re-alignment is squarely within the CPCB methodology"));
children.push(para([
  t("The classification is not left to discretion; it is governed by a national, Pollution-Index (PI) based methodology. The "),
  b("CPCB Directions dated 07-03-2016 (No. B-29012/ESS(CPA)/2015-16)"),
  t(", issued under "),
  b("Section 18(1)(b) of the Water (Prevention and Control of Pollution) Act, 1974"),
  t(", fix the tiers on a Pollution-Index basis — "),
  b("Red (PI \u2265 60), Orange (41\u201359), Green (21\u201340), White (\u2264 20)"),
  t(" — carried forward by the subsequent CPCB Directions of 2025. White-category units require "),
  b("no Consent to Operate"), t("; a mere intimation to the SPCB suffices."),
]));
children.push(para([
  t("Under this very methodology, a "),
  b("dry mechanical de-fibring / ginning operation carries a low Pollution Index"),
  t(" and is treated as White by TNPCB and, in Kerala itself, dry coir manufacturing is already White. Retaining ordinary de-fibring / ginning in Orange in Kerala is therefore "),
  b("stricter than the CPCB norm requires"),
  t(". Retting, which involves water and effluent, legitimately warrants safeguards — but those are best delivered through "),
  b("collective effluent management (Zero-Liquid-Discharge) at a Common Facility Centre"),
  t(", exactly as Tamil Nadu has done, rather than through a Red-category, public-hearing-tier barrier that simply exports the industry and its jobs across the border."),
]));

/* --- 5. Why it matters for Chittur ---------------------------------------- */
children.push(h(5, "The consequence for Chittur"));
children.push(para([
  t("Because of this gap, the "),
  b("primary / value-generating processes (retting, de-fibring, pith buffering) are carried out in Pollachi"),
  t(", while Kerala is left with only downstream value-addition. Coir entrepreneurs — including many from Kerala — have set up their primary units in Tamil Nadu, taking with them the investment, the employment, and the value of "),
  b("our own coconut husk"),
  t(". A narrow, defensible re-alignment of two categories would let that value be retained in Chittur, generating local rural and women's employment in a classic agro-based cottage sector."),
]));

/* --- 6. Prayer ------------------------------------------------------------- */
children.push(h(6, "Our prayer"));
children.push(para([t("We humbly request the Hon'ble Minister, in coordination with the Coir portfolio and the Kerala State Pollution Control Board, to:")]));
children.push(numItem([
  b("Re-align the two bottleneck categories. "),
  t("Reclassify "),
  b("coir de-fibring / ginning from Orange to White"),
  t(" (a dry mechanical process, as Tamil Nadu treats it), and reduce "),
  b("coconut-husk retting from Red to Green"),
  t(" with mandatory Zero-Liquid-Discharge / common effluent safeguards — consistent with the CPCB PI methodology and with TNPCB's White treatment. Bleaching / dyeing / stencilling (which involve chemicals and effluent) may continue to be regulated in the Orange category."),
]));
children.push(numItem([
  b("Grant special coir-cluster / coir industrial-zone status for Chittur, "),
  t("with a Special Purpose Vehicle and a "),
  b("Common Facility Centre"),
  t(" (mechanised de-fibring, pith washing / block-making with Zero-Liquid-Discharge and salinity management), so that environmental concerns are addressed collectively and cost-effectively."),
]));
children.push(numItem([
  b("Extend fiscal incentives — "),
  t("capital subsidy, interest subvention, and "),
  b("power-tariff parity with Tamil Nadu"),
  t(" — to attract coir units to Chittur."),
]));
children.push(numItem([
  b("Ensure a level playing field with Pollachi, "),
  t("so that the value addition to Kerala's own coconut husk is retained within Kerala."),
]));

/* --- 7. Annexures list ----------------------------------------------------- */
children.push(h(7, "Annexures"));
children.push(bullet([b("Annexure A — "), t("KSPCB current classification of coir processes (December 2025).")]));
children.push(bullet([b("Annexure B — "), t("TNPCB: coir retained in the White category — the 10-11-2021 Orange reclassification and its withdrawal on 12-10-2023.")]));
children.push(bullet([b("Annexure C — "), t("Model draft Government Order for the Government's consideration.")]));

/* --- signature block ------------------------------------------------------- */
children.push(new Paragraph({ spacing: { before: 220, after: 60 }, children: [t("Respectfully submitted,")] }));
children.push(new Paragraph({ spacing: { after: 220 }, children: [t("For and on behalf of the Chittur Agriculture Marketing Society, Kozhinjampara")] }));
children.push(new Paragraph({ spacing: { after: 20 }, children: [t("_________________________________")] }));
children.push(new Paragraph({ spacing: { after: 20 }, children: [b("[Name, Designation]")] }));
children.push(new Paragraph({ spacing: { after: 160 }, children: [t("[Phone / Email]")] }));
children.push(new Paragraph({ spacing: { after: 40 }, children: [b("Endorsed by:  "), t("Adv. Sumesh Achuthan, MLA, Chittur.")] }));

/* ========================================================================== */
/* ANNEXURE A                                                                  */
/* ========================================================================== */
children.push(pageBreakPara());
children.push(annexHeading("ANNEXURE A"));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 160 },
  children: [b("KSPCB classification of coir processes — current list (December 2025)")],
}));
children.push(para([
  new TextRun({ text: "Transcribed from the Kerala State Pollution Control Board's current classification of industrial sectors. The source list is to be annexed to this petition in original.", font: "Times New Roman", size: 20, italics: true, color: "555555" }),
]));
{
  const cw = [2000, 5638, 2000]; // = 9638
  const mk = (catText, catColor, activities) => new TableRow({ children: [
    cell([catRun(catText, catColor)], { widthDxa: cw[0], align: AlignmentType.CENTER }),
    cell(activities, { widthDxa: cw[1] }),
    cell(catText === "RED" ? "CTE + CTO; public-hearing-tier possible"
       : catText === "ORANGE" ? "CTE + CTO; periodic monitoring"
       : catText === "GREEN" ? "Simplified consent"
       : "Intimation only (no CTO)", { widthDxa: cw[2] }),
  ]});
  const rows = [
    new TableRow({ tableHeader: true, children: [
      cell("Category", { widthDxa: cw[0], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Coir activity", { widthDxa: cw[1], shade: HEADSHADE, header: true }),
      cell("Consent regime", { widthDxa: cw[2], shade: HEADSHADE, header: true }),
    ]}),
    mk("RED", REDC, "Coconut-husk retting"),
    mk("ORANGE", ORANGEC, "Coir de-fibring / ginning; coir stencilling; coir bleaching / dyeing"),
    mk("GREEN", GREENC, "Coir fibre / pith processing generating effluent"),
    mk("WHITE", WHITEC, "Dry coir processing / coir products without effluent; coir yarn without dyeing; coir-pith final product; coir rope; curled coir"),
  ];
  children.push(table(cw, rows));
}

/* ========================================================================== */
/* ANNEXURE B                                                                  */
/* ========================================================================== */
children.push(pageBreakPara());
children.push(annexHeading("ANNEXURE B"));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 160 },
  children: [b("TNPCB — coir industry retained in the White category")],
}));
children.push(para([
  new TextRun({ text: "Summary of the Tamil Nadu Pollution Control Board's position, with the operative dates and the text of its withdrawal order. Copies of the TNPCB proceedings / press release are to be annexed in original.", font: "Times New Roman", size: 20, italics: true, color: "555555" }),
]));
children.push(bullet([
  b("Baseline position: "),
  t("Under the CPCB Pollution-Index norm (Red high, Orange/Green moderate, White non-polluting), the coconut-husk retting / de-fibering / pith-processing industry stands in the "),
  b("White"), t(" category in Tamil Nadu."),
]));
children.push(bullet([
  b("Proceedings dated 10-11-2021: "),
  t("TNPCB reclassified the coir industry from White to Orange — moving it into the polluting category, requiring TNPCB approval to set up units and subjecting them to periodic monitoring — stated to be pursuant to directions of the environmental tribunal / courts."),
]));
children.push(bullet([
  b("Withdrawal dated 12-10-2023: "),
  t("Following an enormous number of representations from various associations of coir exporters and manufacturers seeking protection of the coir sector and their livelihood, TNPCB withdrew the Orange reclassification. The operative reasoning recorded was:"),
]));
children.push(new Paragraph({
  spacing: { before: 40, after: 120, line: 276 },
  indent: { left: 480, right: 360 },
  border: { left: { color: RULE, style: BorderStyle.SINGLE, size: 12, space: 8 } },
  children: [new TextRun({
    text: "\u201CConsidering the coir industry is one of the vital agro-based cottage-type industries in the State contributing significantly to the creation of livelihood in major coconut-growing districts and to encourage the cottage-type MSME sector for the production of eco-friendly products, TNPCB withdraws the proceeding regarding the categorisation of coconut husk retting / de-fibres / pith processing under the Orange category.\u201D",
    font: "Times New Roman", size: 21, italics: true, color: "333333",
  })],
}));
children.push(bullet([
  b("Effect: "),
  t("The coconut-husk retting / de-fibring / pith-processing industry "),
  b("remains in the White category in Tamil Nadu"),
  t(", relieving these cottage/MSME units of Consent-to-Operate and monitoring burdens."),
]));
children.push(para([
  b("Relevance to this petition: "),
  t("Kerala keeps husk retting in Red and de-fibring / ginning in Orange for the "),
  b("same raw material and the same process chain"),
  t(" that Tamil Nadu treats as White, on the other side of the same district border. That contrast is the core justification for the re-alignment prayed for in paragraph 6."),
]));

/* ========================================================================== */
/* ANNEXURE C — model draft G.O.                                               */
/* ========================================================================== */
children.push(pageBreakPara());
children.push(annexHeading("ANNEXURE C"));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [b("Model draft Government Order (for the Government's consideration)")],
}));
children.push(para([
  new TextRun({ text: "The following is a model draft placed for the Government's consideration. The actual G.O. number, date and file references are to be assigned by the Department.", font: "Times New Roman", size: 20, italics: true, color: "555555" }),
]));

const goLines = [
  "GOVERNMENT OF KERALA",
  "Abstract",
  "",
  "Industries Department — Coir sector in Chittur (Palakkad) — Re-alignment of the",
  "pollution-category of coconut-husk retting (Red \u2192 Green with ZLD) and coir",
  "de-fibring / ginning (Orange \u2192 White) with the CPCB Pollution-Index Directions",
  "and with the Tamil Nadu White-category treatment, and grant of special coir-cluster",
  "status with a Common Facility Centre — Orders issued.",
  "",
  "INDUSTRIES (COIR) DEPARTMENT",
  "G.O.(Ms) No. ______/2026/ID            Thiruvananthapuram, dated __-__-2026",
  "",
  "Read: 1. CPCB Directions dated 07-03-2016 (No. B-29012/ESS(CPA)/2015-16) u/s",
  "         18(1)(b) of the Water (Prevention and Control of Pollution) Act, 1974,",
  "         and the subsequent CPCB Directions of 2025.",
  "      2. TNPCB proceedings dated 10-11-2021 and their withdrawal dated 12-10-2023.",
  "      3. Representation dated __-07-2026 of the Chittur Agriculture Marketing",
  "         Society, Kozhinjampara, forwarded by the MLA, Chittur.",
  "",
  "ORDER",
  "",
  "  Whereas coconut-husk retting and coir de-fibring / ginning are treated more",
  "stringently in Kerala (Red and Orange respectively) than the CPCB Pollution-Index",
  "norm requires, and than Tamil Nadu treats the identical process chain (White),",
  "causing these units and their employment to migrate to Pollachi (Tamil Nadu);",
  "",
  "  Government, having considered the matter, hereby order that —",
  "",
  "  (1) the Kerala State Pollution Control Board shall re-align coir de-fibring /",
  "      ginning to the White category and coconut-husk retting to the Green category",
  "      with mandatory Zero-Liquid-Discharge safeguards, consistent with the CPCB",
  "      Directions read above, retaining Orange regulation only for bleaching /",
  "      dyeing / stencilling;",
  "  (2) a special coir-cluster / coir industrial-zone status is accorded to Chittur,",
  "      with an SPV and a Common Facility Centre (mechanised de-fibring, pith",
  "      washing / block-making with ZLD and salinity management);",
  "  (3) capital subsidy, interest subvention and power-tariff parity with Tamil Nadu",
  "      shall be extended to coir units in Chittur; and",
  "  (4) the Industries and Coir Directorates shall ensure a level playing field so",
  "      that value addition to Kerala's coconut husk is retained within the State.",
  "",
  "                                        By order of the Governor,",
  "                                        __________________________",
  "                                        Secretary to Government",
  "",
  "To  The Director of Industries & Commerce; the Director of Coir Development;",
  "    the Chairman, Kerala State Pollution Control Board.",
  "    Copy to: the Minister for Coir; Environment Department; the MLA, Chittur.",
];
{
  const goPara = new Paragraph({
    spacing: { before: 40, after: 40, line: 240 },
    children: goLines.flatMap((ln, i) => {
      const run = new TextRun({ text: ln || " ", font: "Consolas", size: 17, color: INK });
      return i === 0 ? [run] : [new TextRun({ break: 1, text: ln || " ", font: "Consolas", size: 17, color: INK })];
    }),
  });
  const goCell = new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    shading: { type: ShadingType.CLEAR, fill: "F5F5F0", color: "auto" },
    children: [goPara],
  });
  children.push(new Table({
    columnWidths: [CONTENT_W],
    width: { size: CONTENT_W, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      left: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
    },
    rows: [new TableRow({ children: [goCell] })],
  }));
}

/* ========================================================================== */
/* DOCUMENT                                                                     */
/* ========================================================================== */
const doc = new Document({
  creator: "Chittur Agriculture Marketing Society, Kozhinjampara",
  title: "Petition — Coir Industry Pollution-Category Parity, Chittur",
  description: "Formal petition on KSPCB vs TNPCB coir classification parity for Chittur.",
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 22, color: INK } },
    },
    paragraphStyles: [{
      id: "Heading1",
      name: "Heading 1",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      run: { font: "Times New Roman", size: 24, bold: true, color: GREEN },
      paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 0 },
    }],
  },
  numbering: {
    config: [
      {
        reference: "petition-list",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 620, hanging: 340 } } } }],
      },
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 520, hanging: 260 } } } }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1276, bottom: 1134, left: 1134, right: 1134 },
      },
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 0 },
        border: { bottom: { color: "CCCCCC", style: BorderStyle.SINGLE, size: 4, space: 2 } },
        children: [new TextRun({ text: "Chittur Agriculture Marketing Society, Kozhinjampara", font: "Times New Roman", size: 16, italics: true, color: "888888" })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          new TextRun({ text: "Coir pollution-category parity petition \u2014 Chittur", font: "Times New Roman", size: 16, color: "999999" }),
          new TextRun({ text: "\tPage ", font: "Times New Roman", size: 16, color: "999999" }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 16, color: "999999" }),
          new TextRun({ text: " of ", font: "Times New Roman", size: 16, color: "999999" }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Times New Roman", size: 16, color: "999999" }),
        ],
      })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = "Coir-Industry-Pollution-Category-Petition.docx";
  fs.writeFileSync(out, buf);
  console.log("Wrote", out, buf.length, "bytes");
});
