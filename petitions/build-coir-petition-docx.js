/*
 * Formal petition (Word .docx) — Pollution-category parity for coir processing in Chittur.
 * Petitioner: the coconut farmers of the Chittur Assembly Constituency, represented by Arun.
 * Addressed to the Minister in charge of Pollution Control (Environment), Government of Kerala;
 * copies to the Chief Minister and the Minister for Industries & Commerce.
 *
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
  ShadingType, BorderStyle, PageNumber, TabStopType,
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
const FLAGC = "8A3324";        // divergence flag colour

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

function spacer(after = 120) {
  return new Paragraph({ spacing: { after }, children: [t("")] });
}

/* ---- table helpers (dual widths, CLEAR shading) --------------------------- */
function cell(children, { widthDxa, shade, header = false, align = AlignmentType.LEFT, size = 20 } = {}) {
  const kids = (Array.isArray(children) ? children : [children]).map((c) =>
    typeof c === "string"
      ? new Paragraph({
          alignment: align,
          spacing: { after: 30, line: 248 },
          children: [new TextRun({ text: c, font: "Times New Roman", size, bold: header, color: header ? "222222" : INK })],
        })
      : c
  );
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: "auto" } : undefined,
    children: kids,
  });
}

function catRun(text, color, align = AlignmentType.CENTER) {
  return new Paragraph({
    alignment: align,
    spacing: { after: 30, line: 248 },
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
  children: [new TextRun({ text: "PETITION OF THE COCONUT FARMERS OF CHITTUR ASSEMBLY CONSTITUENCY", font: "Times New Roman", size: 26, bold: true, color: INK })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 40 },
  children: [new TextRun({ text: "Palakkad District, Kerala \u2014 represented by Arun", font: "Times New Roman", size: 20, italics: true, color: "555555" })],
}));
children.push(thinRule());

/* --- title ---------------------------------------------------------------- */
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 60, after: 60 },
  children: [new TextRun({ text: "POLLUTION-CATEGORY PARITY FOR COIR PROCESSING IN CHITTUR", font: "Times New Roman", size: 26, bold: true, color: INK })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 180 },
  children: [new TextRun({ text: "Why Pollachi (Tamil Nadu) became a coir-processing hub while Chittur \u2014 with the same husk, 80 km away \u2014 did not", font: "Times New Roman", size: 21, italics: true, color: "444444" })],
}));

/* --- addressing block ----------------------------------------------------- */
children.push(labelLine("To:", [
  t("The Hon'ble Minister in charge of Pollution Control (Environment), Government of Kerala,"), new TextRun({ break: 1 }),
  t("Thiruvananthapuram."), new TextRun({ break: 1 }),
  new TextRun({ text: "(Administrative Department of the Kerala State Pollution Control Board.)", font: "Times New Roman", size: 20, italics: true, color: "555555" }),
]));
children.push(labelLine("Copy to:", [
  t("1. The Hon'ble Chief Minister of Kerala."), new TextRun({ break: 1 }),
  t("2. The Hon'ble Minister for Industries & Commerce, Government of Kerala."),
]));
children.push(labelLine("From:", [
  t("The coconut farmers of the Chittur Assembly Constituency, Palakkad District, Kerala \u2014 represented by "),
  b("Arun"), t(" and delegation."),
]));
children.push(labelLine("Date:", [b("[__] July 2026"), t(", Thiruvananthapuram.")]));
children.push(new Paragraph({
  spacing: { before: 40, after: 160, line: 276 },
  children: [
    b("Subject:  "),
    t("Request to re-align the pollution-control category of coconut-husk retting (presently Red) and coir de-fibring / ginning (presently Orange) in Kerala with the Central (CPCB) norms and with Tamil Nadu's White-category treatment of the identical process chain \u2014 so that coir processing develops in Chittur instead of migrating across the border to Pollachi."),
  ],
}));
children.push(thinRule());

/* --- 1. The ask, in one line ---------------------------------------------- */
children.push(h(1, "The prayer in one line"));
children.push(para([
  t("We are coconut farmers of the Chittur Assembly Constituency. We grow the coconut; the husk that comes off it is coir's raw material. We ask the Government to fix a single, provable anomaly: "),
  b("Kerala places coconut-husk retting in the Red category and coir de-fibring / ginning in the Orange category, while both the Central norm and neighbouring Tamil Nadu treat the very same activities as practically non-polluting (White)."),
  t(" That one gap has sent the industry \u2014 and our husk's value \u2014 across the border to Pollachi."),
]));

/* --- 2. The three-way comparison table ------------------------------------ */
children.push(h(2, "The classification gap \u2014 Kerala vs. Tamil Nadu vs. the Central (CPCB) norm"));
children.push(para([
  t("The table below sets each coir activity against its category in Kerala (KSPCB), Tamil Nadu (TNPCB) and the Central norm (CPCB). The divergence is confined to the "),
  b("primary end"), t(" of the chain:"),
]));
{
  const cw = [2760, 1180, 1180, 1560, 2958]; // = 9638
  const catCell = (txt, color, w) => cell([catRun(txt, color)], { widthDxa: w });
  const rows = [
    new TableRow({ tableHeader: true, children: [
      cell("Coir activity", { widthDxa: cw[0], shade: HEADSHADE, header: true }),
      cell("Kerala (KSPCB)", { widthDxa: cw[1], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Tamil Nadu (TNPCB)", { widthDxa: cw[2], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Central (CPCB)", { widthDxa: cw[3], shade: HEADSHADE, header: true, align: AlignmentType.CENTER }),
      cell("Where the outlier is", { widthDxa: cw[4], shade: HEADSHADE, header: true }),
    ]}),
    new TableRow({ children: [
      cell("Coconut-husk retting", { widthDxa: cw[0] }),
      catCell("RED", REDC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      cell([catRun("Not a Red / Orange sector \u2020", GREENC)], { widthDxa: cw[3] }),
      cell([catRun("Kerala alone \u2014 Red vs. White elsewhere", FLAGC, AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
    new TableRow({ children: [
      cell("Coir de-fibring / ginning", { widthDxa: cw[0] }),
      catCell("ORANGE", ORANGEC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      cell([catRun("White / Green (dry) \u2020", GREENC)], { widthDxa: cw[3] }),
      cell([catRun("Kerala alone \u2014 Orange vs. White elsewhere", FLAGC, AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
    new TableRow({ children: [
      cell("Coir fibre / pith processing generating effluent", { widthDxa: cw[0] }),
      catCell("GREEN", GREENC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      catCell("GREEN", GREENC, cw[3]),
      cell([catRun("Kerala & CPCB stricter than TN", "555555", AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
    new TableRow({ children: [
      cell("Coir bleaching / dyeing / stencilling (chemical, wet)", { widthDxa: cw[0] }),
      catCell("ORANGE", ORANGEC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      catCell("ORANGE", ORANGEC, cw[3]),
      cell([catRun("Kerala aligned with CPCB", "555555", AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
    new TableRow({ children: [
      cell("Dry coir processing / coir products without effluent", { widthDxa: cw[0] }),
      catCell("WHITE", WHITEC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      catCell("WHITE", WHITEC, cw[3]),
      cell([catRun("Fully aligned", "555555", AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
    new TableRow({ children: [
      cell("Coir yarn without dyeing", { widthDxa: cw[0] }),
      catCell("WHITE", WHITEC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      catCell("WHITE", WHITEC, cw[3]),
      cell([catRun("Fully aligned", "555555", AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
    new TableRow({ children: [
      cell("Coir-pith final product; coir rope; curled coir", { widthDxa: cw[0] }),
      catCell("WHITE", WHITEC, cw[1]),
      catCell("WHITE", WHITEC, cw[2]),
      catCell("WHITE", WHITEC, cw[3]),
      cell([catRun("Fully aligned", "555555", AlignmentType.LEFT)], { widthDxa: cw[4] }),
    ]}),
  ];
  children.push(table(cw, rows));
}
children.push(new Paragraph({
  spacing: { before: 60, after: 120, line: 252 },
  children: [new TextRun({
    text: "\u2020 CPCB's 2016 categorisation list places \u201CManufacturing of coir items from coconut husks\u201D in the White category; its 2023 draft re-categorisation places dry coir processing in Green, coir items in White, and only wet / dyeing coir in Orange. CPCB does not list ordinary husk retting or dry de-fibring as a Red sector.",
    font: "Times New Roman", size: 18, italics: true, color: "555555",
  })],
}));
children.push(para([
  t("Read down the last column: on "),
  b("five of seven"), t(" activities the three regulators broadly agree. The entire dispute is the "),
  b("top two rows"), t(" \u2014 retting and de-fibring / ginning \u2014 where "),
  b("Kerala stands alone"), t(" in imposing Red / Orange treatment on activities that Tamil Nadu and the Central norm regard as White."),
]));

/* --- 3. Why Pollachi is the hub ------------------------------------------- */
children.push(h(3, "Why this made Pollachi the hub and left Chittur out"));
children.push(para([
  t("Category is not a label \u2014 it decides whether a small unit can exist. A "),
  b("White"), t(" activity needs "),
  b("no Consent to Operate"), t("; a simple intimation suffices, and the unit can start almost at once. A "),
  b("Red"), t(" activity needs a full Consent to Establish and Consent to Operate and can attract "),
  b("public-hearing-tier scrutiny"), t("; an "),
  b("Orange"), t(" activity needs the same consents plus periodic monitoring. For a cottage / MSME coir unit, that is the difference between opening and not opening."),
]));
children.push(para([
  t("So a coir entrepreneur \u2014 including many from Kerala \u2014 does the obvious thing: he sets up retting and de-fibring "),
  b("80 km away in Pollachi"), t(", where they are White, and leaves Chittur out. The result on the ground:"),
]));
children.push(bullet([b("The husk moves, not the value. "), t("Our coconut husk is carted to Pollachi, processed there, and the fibre, pith, jobs and margins are booked in Tamil Nadu.")]));
children.push(bullet([b("Pollachi is a coir cluster; Chittur is a supplier of raw husk. "), t("The same raw material and the same process chain, but the primary industry clustered on the side of the border where the category is lighter.")]));
children.push(bullet([b("Farmers lose twice. "), t("We forgo the price a local processing demand would give our husk, and we forgo the local employment that retting, de-fibring and pith units would create in the Chittur belt.")]));

/* --- 4. Legal basis -------------------------------------------------------- */
children.push(h(4, "The re-alignment is squarely within the Central norm \u2014 not a dilution of safeguards"));
children.push(para([
  t("Classification is governed by a national, Pollution-Index (PI) methodology. The "),
  b("CPCB Directions dated 07-03-2016 (No. B-29012/ESS(CPA)/2015-16)"),
  t(", under "),
  b("Section 18(1)(b) of the Water (Prevention and Control of Pollution) Act, 1974"),
  t(", fix the tiers by PI \u2014 "),
  b("Red \u2265 60, Orange 41\u201359, Green 21\u201340, White \u2264 20"),
  t(" \u2014 carried forward by the CPCB Directions of 2025. On that very scale a dry mechanical "),
  b("de-fibring / ginning"), t(" operation is low-index and is treated as White; the Central list itself places coir manufacturing from coconut husks in White."),
]));
children.push(para([
  t("Tamil Nadu proved the point in practice. By proceedings dated "),
  b("10-11-2021"), t(" TNPCB moved coir from White to Orange; after representations from coir associations it "),
  b("withdrew"), t(" that on "),
  b("12-10-2023"), t(", restoring the \u201Ccoconut-husk retting / de-fibering / pith-processing industry\u201D to White, expressly to protect the agro-based cottage MSME sector (Annexure B). Retting does use water \u2014 but that is best handled by "),
  b("Zero-Liquid-Discharge at a shared Common Facility Centre"), t(", exactly as Tamil Nadu does, not by a Red-category barrier that simply exports the industry."),
]));

/* --- 5. Prayer ------------------------------------------------------------- */
children.push(h(5, "Our prayer"));
children.push(para([t("We humbly request the Hon'ble Minister in charge of Pollution Control, in consultation with the Chief Minister and the Minister for Industries, to direct the Kerala State Pollution Control Board to:")]));
children.push(numItem([
  b("Reclassify coir de-fibring / ginning from Orange to White"),
  t(" (a dry mechanical process, as the Central norm and Tamil Nadu treat it)."),
]));
children.push(numItem([
  b("Reduce coconut-husk retting from Red to Green"),
  t(", with mandatory Zero-Liquid-Discharge / common effluent safeguards, in line with the CPCB PI methodology and Tamil Nadu's White treatment. Bleaching / dyeing / stencilling may remain regulated (Orange), as they alone involve chemicals."),
]));
children.push(numItem([
  b("Sanction a coir Common Facility Centre and cluster status for Chittur"),
  t(" (mechanised de-fibring, pith washing / block-making with ZLD and salinity management), so that environmental safeguards are met collectively and cheaply."),
]));
children.push(numItem([
  b("Ensure a level playing field with Pollachi"),
  t(" \u2014 including power-tariff parity and MSME incentives \u2014 so that the value of Kerala's own coconut husk is retained in Kerala."),
]));

/* --- 6. Annexures list ----------------------------------------------------- */
children.push(h(6, "Annexures"));
children.push(bullet([b("Annexure A \u2014 "), t("KSPCB current classification of coir processes (December 2025).")]));
children.push(bullet([b("Annexure B \u2014 "), t("TNPCB: coir retained in the White category \u2014 the 10-11-2021 Orange reclassification and its withdrawal on 12-10-2023.")]));
children.push(bullet([b("Annexure C \u2014 "), t("Model draft Government Order for the Government's consideration.")]));

/* --- signature block ------------------------------------------------------- */
children.push(new Paragraph({ spacing: { before: 220, after: 60 }, children: [t("Respectfully submitted,")] }));
children.push(new Paragraph({ spacing: { after: 220 }, children: [t("For and on behalf of the coconut farmers of the Chittur Assembly Constituency,")] }));
children.push(new Paragraph({ spacing: { after: 20 }, children: [t("_________________________________")] }));
children.push(new Paragraph({ spacing: { after: 20 }, children: [b("Arun"), t("  (authorised representative of the petitioner farmers)")] }));
children.push(new Paragraph({ spacing: { after: 60 }, children: [t("[Phone / Email]")] }));

/* ========================================================================== */
/* ANNEXURE A                                                                  */
/* ========================================================================== */
children.push(pageBreakPara());
children.push(annexHeading("ANNEXURE A"));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 160 },
  children: [b("KSPCB classification of coir processes \u2014 current list (December 2025)")],
}));
children.push(para([
  new TextRun({ text: "Transcribed from the Kerala State Pollution Control Board's current classification of industrial sectors. The source list is to be annexed to this petition in original.", font: "Times New Roman", size: 20, italics: true, color: "555555" }),
]));
{
  const cw = [2000, 5638, 2000]; // = 9638
  const mk = (catText, catColor, activities) => new TableRow({ children: [
    cell([catRun(catText, catColor)], { widthDxa: cw[0] }),
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
  children: [b("TNPCB \u2014 coir industry retained in the White category")],
}));
children.push(para([
  new TextRun({ text: "Summary of the Tamil Nadu Pollution Control Board's position, with the operative dates and the text of its withdrawal order. Copies of the TNPCB proceedings / press release are to be annexed in original.", font: "Times New Roman", size: 20, italics: true, color: "555555" }),
]));
children.push(bullet([
  b("Baseline position: "),
  t("Under the CPCB Pollution-Index norm (Red high, Orange / Green moderate, White non-polluting), the coconut-husk retting / de-fibering / pith-processing industry stands in the "),
  b("White"), t(" category in Tamil Nadu."),
]));
children.push(bullet([
  b("Proceedings dated 10-11-2021: "),
  t("TNPCB reclassified the coir industry from White to Orange \u2014 moving it into the polluting category, requiring TNPCB approval to set up units and subjecting them to periodic monitoring \u2014 stated to be pursuant to directions of the environmental tribunal / courts."),
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
  t(", relieving these cottage / MSME units of Consent-to-Operate and monitoring burdens."),
]));
children.push(para([
  b("Relevance to this petition: "),
  t("Kerala keeps husk retting in Red and de-fibring / ginning in Orange for the "),
  b("same raw material and the same process chain"),
  t(" that both Tamil Nadu and the Central norm treat as White, on the other side of the same district border. That contrast is the core justification for the re-alignment prayed for in paragraph 5."),
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
  "Environment Department \u2014 Coir sector in Chittur (Palakkad) \u2014 Re-alignment of the",
  "pollution-category of coconut-husk retting (Red \u2192 Green with ZLD) and coir",
  "de-fibring / ginning (Orange \u2192 White) with the CPCB Pollution-Index Directions",
  "and with the Tamil Nadu White-category treatment, and grant of a coir Common",
  "Facility Centre with cluster status \u2014 Orders issued.",
  "",
  "ENVIRONMENT (POLLUTION CONTROL) DEPARTMENT",
  "G.O.(Ms) No. ______/2026/Envt.        Thiruvananthapuram, dated __-__-2026",
  "",
  "Read: 1. CPCB Directions dated 07-03-2016 (No. B-29012/ESS(CPA)/2015-16) u/s",
  "         18(1)(b) of the Water (Prevention and Control of Pollution) Act, 1974,",
  "         and the subsequent CPCB Directions of 2025.",
  "      2. TNPCB proceedings dated 10-11-2021 and their withdrawal dated 12-10-2023.",
  "      3. Representation dated __-07-2026 of the coconut farmers of the Chittur",
  "         Assembly Constituency, represented by Arun.",
  "",
  "ORDER",
  "",
  "  Whereas coconut-husk retting and coir de-fibring / ginning are treated more",
  "stringently in Kerala (Red and Orange respectively) than the CPCB Pollution-Index",
  "norm requires, and than Tamil Nadu treats the identical process chain (White),",
  "causing these units and their employment to migrate to Pollachi (Tamil Nadu);",
  "",
  "  Government, having considered the matter, hereby order that \u2014",
  "",
  "  (1) the Kerala State Pollution Control Board shall re-align coir de-fibring /",
  "      ginning to the White category and coconut-husk retting to the Green category",
  "      with mandatory Zero-Liquid-Discharge safeguards, consistent with the CPCB",
  "      Directions read above, retaining Orange regulation only for bleaching /",
  "      dyeing / stencilling;",
  "  (2) a coir Common Facility Centre with cluster status is sanctioned for Chittur",
  "      (mechanised de-fibring, pith washing / block-making with ZLD and salinity",
  "      management);",
  "  (3) power-tariff parity with Tamil Nadu and MSME incentives shall be extended to",
  "      coir units in Chittur; and",
  "  (4) the Environment, Industries and Coir Directorates shall ensure a level",
  "      playing field so that value addition to Kerala's coconut husk is retained",
  "      within the State.",
  "",
  "                                        By order of the Governor,",
  "                                        __________________________",
  "                                        Secretary to Government",
  "",
  "To  The Chairman, Kerala State Pollution Control Board; the Director of Industries",
  "    & Commerce; the Director of Coir Development.",
  "    Copy to: the Chief Minister's Office; the Minister for Industries; the MLA, Chittur.",
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
  creator: "Coconut farmers of Chittur Assembly Constituency (represented by Arun)",
  title: "Petition — Coir Processing Pollution-Category Parity, Chittur",
  description: "Petition by Chittur coconut farmers on KSPCB vs TNPCB vs CPCB coir classification parity.",
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
        children: [new TextRun({ text: "Coconut farmers of Chittur Assembly Constituency", font: "Times New Roman", size: 16, italics: true, color: "888888" })],
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
