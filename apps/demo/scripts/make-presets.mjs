import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  Math,
  MathFraction,
  MathIntegral,
  MathRadical,
  MathRun,
  MathSum,
  MathSuperScript,
  PageBreak,
  PageNumber,
  Packer,
  Paragraph,
  SectionType,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const outDir = join(dirname(fileURLToPath(import.meta.url)), "../public/fixtures");
mkdirSync(outDir, { recursive: true });

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = {
  top: noBorder,
  bottom: noBorder,
  left: noBorder,
  right: noBorder,
  insideHorizontal: noBorder,
  insideVertical: noBorder,
};
const lightBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
};

const formula = (children) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 110, after: 110 },
  children: [new Math({ children })],
});

const equations = new Document({
  styles: { default: { document: { run: { font: "Cambria", size: 22, color: "243447" } } } },
  sections: [{
    properties: { page: { margin: { top: 850, right: 900, bottom: 850, left: 900 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Mathematical methods  •  ", PageNumber.CURRENT], color: "64748B", size: 18 })] })] }) },
    children: [
      new Paragraph({ children: [new TextRun({ text: "A compact equation sampler", bold: true, size: 38, color: "17365D" })] }),
      new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Native Office Math structures remain editable and scale with the surrounding text.", color: "526579" })] }),
      new Paragraph({ text: "Energy and wavelength", heading: HeadingLevel.HEADING_2 }),
      formula([new MathRun("E = "), new MathFraction({ numerator: [new MathRun("hc")], denominator: [new MathRun("λ")] })]),
      new Paragraph({ text: "Euler’s identity and a radical", heading: HeadingLevel.HEADING_2 }),
      formula([
        new MathSuperScript({ children: [new MathRun("e")], superScript: [new MathRun("iπ")] }),
        new MathRun(" + 1 = 0     "),
        new MathRadical({ degree: [new MathRun("3")], children: [new MathRun("x + 1")] }),
      ]),
      new Paragraph({ text: "Integral and weighted mean", heading: HeadingLevel.HEADING_2 }),
      formula([
        new MathIntegral({
          children: [new MathSuperScript({ children: [new MathRun("e")], superScript: [new MathRun("−t²")] }), new MathRun("dt")],
          subScript: [new MathRun("0")],
          superScript: [new MathRun("∞")],
        }),
        new MathRun(" = "),
        new MathFraction({ numerator: [new MathRadical({ children: [new MathRun("π")] })], denominator: [new MathRun("2")] }),
      ]),
      formula([
        new MathRun("μ = "),
        new MathFraction({
          numerator: [new MathSum({ children: [new MathRun("wᵢxᵢ")], subScript: [new MathRun("i=1")], superScript: [new MathRun("n")] })],
          denominator: [new MathSum({ children: [new MathRun("wᵢ")], subScript: [new MathRun("i=1")], superScript: [new MathRun("n")] })],
        }),
      ]),
    ],
  }],
});

const tableHeader = (text) => new TableCell({
  shading: { type: ShadingType.CLEAR, fill: "17365D" },
  children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF" })] })],
});

const tables = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 20, color: "26364A" } } } },
  sections: [{
    properties: { page: { margin: { top: 700, right: 700, bottom: 700, left: 700 } } },
    children: [
      new Paragraph({ children: [new TextRun({ text: "Quarterly operating review", bold: true, size: 38, color: "17365D" })] }),
      new Paragraph({ spacing: { after: 180 }, children: [new TextRun({ text: "Tables, merged cells, fills, alignment, borders, and wrapping.", color: "526579" })] }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: lightBorders,
        rows: [
          new TableRow({ tableHeader: true, children: ["Region", "Revenue", "Change", "Status"].map(tableHeader) }),
          new TableRow({ children: ["North America", "$4.2M", "+12%", "Ahead"].map((text) => new TableCell({ children: [new Paragraph(text)] })) }),
          new TableRow({ children: ["Europe", "$2.8M", "+7%", "On plan"].map((text) => new TableCell({ children: [new Paragraph(text)] })) }),
          new TableRow({ children: ["Asia Pacific", "$1.9M", "+16%", "Ahead"].map((text) => new TableCell({ children: [new Paragraph(text)] })) }),
          new TableRow({ children: [
            new TableCell({ columnSpan: 2, shading: { type: ShadingType.CLEAR, fill: "EAF1F8" }, children: [new Paragraph({ children: [new TextRun({ text: "Total revenue", bold: true })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "EAF1F8" }, children: [new Paragraph({ children: [new TextRun({ text: "$8.9M", bold: true })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "EAF1F8" }, children: [new Paragraph({ children: [new TextRun({ text: "+11%", bold: true })] })] }),
          ] }),
        ],
      }),
      new Paragraph({ spacing: { before: 250, after: 80 }, children: [new TextRun({ text: "Milestone plan", bold: true, size: 26, color: "1E3A5F" })] }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [1700, 5100, 1800],
        borders: lightBorders,
        rows: [
          new TableRow({ tableHeader: true, children: ["Team", "Deliverable", "Due"].map(tableHeader) }),
          new TableRow({ children: [new TableCell({ children: [new Paragraph("Documentation")] }), new TableCell({ children: [new Paragraph("Publish the customer migration guide and support playbook")]}), new TableCell({ children: [new Paragraph("Aug 02")] })] }),
          new TableRow({ children: [new TableCell({ children: [new Paragraph("Design")] }), new TableCell({ children: [new Paragraph("Complete the billing dashboard usability review")]}), new TableCell({ children: [new Paragraph("Aug 09")] })] }),
          new TableRow({ children: [new TableCell({ children: [new Paragraph("Operations")] }), new TableCell({ children: [new Paragraph("Finalize launch readiness across all regions")]}), new TableCell({ children: [new Paragraph("Aug 16")] })] }),
        ],
      }),
    ],
  }],
});

const storyParagraphs = [
  "A once-quiet stretch of waterfront is becoming a living laboratory for cleaner streets, cooler buildings, and public spaces that work harder for the people around them.",
  "The change is easiest to see before sunrise. Delivery vans now gather at a shared hub, where parcels are transferred to cargo bikes for the last mile. By eight o’clock, the curb is clear for buses, café tables, and school drop-off.",
  "The project began with a modest question: what if a neighborhood treated energy, mobility, and shade as one connected system? Planners paired heat maps with resident interviews, then tested each idea one block at a time.",
  "Early results are encouraging. Summer surface temperatures have fallen near newly planted corridors, while local merchants report steadier foot traffic throughout the day.",
  "Skeptics remain, and some of their questions are hard ones. Who maintains the new plantings in a drought year? What happens to delivery costs when the pilot subsidies end? The project office publishes its budget and sensor data monthly, a level of openness that has quieted some critics and armed others.",
  "Other cities are watching. Delegations have toured the district in growing numbers, and the planning office now runs a standing workshop on what transfers elsewhere and what depends on local geography.",
  "The honest answer, staff say, is that no single piece is novel. Shade trees, freight hubs, and slow streets are old ideas. The experiment is in running them together, measuring everything, and being willing to unwind what fails.",
  "The work is not finished. The next phase will add rain gardens, retrofit older apartments, and extend the low-traffic network toward the regional rail station.",
  "For longtime residents, the most important measure is less technical: the waterfront feels like a place to linger again.",
];

const publication = new Document({
  styles: { default: { document: { run: { font: "Georgia", size: 20, color: "273444" } } } },
  sections: [
    {
      properties: { type: SectionType.CONTINUOUS, page: { margin: { top: 620, right: 700, bottom: 650, left: 700 } } },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "THE SUNDAY REVIEW", bold: true, size: 19, color: "9A3412", characterSpacing: 140 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 50 }, children: [new TextRun({ text: "How one neighborhood is designing a cooler future", bold: true, size: 43, color: "172A3A" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 150 }, children: [new TextRun({ text: "Small urban experiments are adding up to a new model for city life.", italics: true, color: "526579", size: 23 })] }),
      ],
    },
    {
      // CONTINUOUS: the column section flows on the same page as the title;
      // the default (new page) stranded the title alone with a page of
      // white space and pushed the story to page 2.
      properties: { type: SectionType.CONTINUOUS, page: { margin: { top: 620, right: 700, bottom: 650, left: 700 } }, column: { count: 2, space: 420, separate: true } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Sunday Review  •  ", PageNumber.CURRENT], size: 18, color: "64748B" })] })] }) },
      children: [
        new Paragraph({ spacing: { after: 90 }, children: [new TextRun({ text: "SPECIAL REPORT", bold: true, size: 18, color: "9A3412" })] }),
        ...storyParagraphs.map((text, index) => new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 120, line: 260 },
          children: index === 0
            ? [new TextRun({ text: text[0], bold: true, size: 42, color: "9A3412" }), new TextRun(text.slice(1))]
            : [new TextRun(text)],
        })),
        new Paragraph({
          spacing: { before: 150, after: 80 },
          border: { top: { style: BorderStyle.SINGLE, size: 8, space: 8, color: "9A3412" } },
          children: [new TextRun({ text: "BY THE NUMBERS", bold: true, color: "9A3412" })],
        }),
        new Paragraph({ children: [new TextRun({ text: "18%", bold: true, size: 34 }), new TextRun(" reduction in peak surface temperature along shaded blocks.")] }),
      ],
    },
    {
      // A trailing CONTINUOUS single-column section: Word only BALANCES the
      // two columns above when another continuous break follows — without
      // this the whole story stacks in the left column and the right one
      // stays empty.
      properties: { type: SectionType.CONTINUOUS, page: { margin: { top: 620, right: 700, bottom: 650, left: 700 } } },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 220 }, children: [new TextRun({ text: "◆  ◆  ◆", color: "9A3412", size: 18 })] }),
      ],
    },
  ],
});

const chapterBook = new Document({
  evenAndOddHeaderAndFooters: true,
  styles: {
    default: { document: { run: { font: "Garamond", size: 23, color: "2D2926" }, paragraph: { spacing: { line: 320 } } } },
  },
  sections: [{
    properties: { page: { margin: { top: 1050, right: 1100, bottom: 950, left: 1100 } }, titlePage: true },
    headers: {
      default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "THE MAP OF QUIET THINGS", size: 17, color: "7C6F64" })] })] }),
      even: new Header({ children: [new Paragraph({ children: [new TextRun({ text: "CHAPTER ONE", size: 17, color: "7C6F64" })] })] }),
      first: new Header({ children: [new Paragraph("")] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "7C6F64" })] })] }),
    },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 900, after: 140 }, children: [new TextRun({ text: "CHAPTER ONE", bold: true, size: 19, color: "8B5E3C", characterSpacing: 120 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 460 }, children: [new TextRun({ text: "The Room Above the Station", bold: true, size: 39 })] }),
      new Paragraph({ spacing: { after: 180 }, children: [new TextRun({ text: "T", bold: true, size: 48, color: "8B5E3C" }), new TextRun("he room had been empty for eleven years, though no one at the station could agree on what it had once contained.")] }),
      new Paragraph("Mara found the key beneath a ledger in the ticket office. It was small and dark with age, tied to a paper label that read simply: ABOVE."),
      new Paragraph("At the end of the evening shift, after the last train had carried its bright windows north, she climbed the narrow stair beside platform three."),
      new Paragraph("The key turned without resistance. Beyond the door, moonlight rested on rows of shallow drawers, each fitted with a brass pull and a handwritten name."),
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph("She opened the nearest drawer. Inside was a paper packet, a loop of blue thread, and the faint smell of rain on warm pavement."),
      new Paragraph("The label read JULY 14, 1998 — THE SOUND OF THE 6:10 LEAVING. Mara glanced toward the window. Far below, the rails shone like two lines of ink."),
      new Paragraph("Another drawer held the hush that falls just before snow. A third contained the click of her mother’s garden gate. The room, she realized, was an archive of things too ordinary to save and too important to lose."),
      new Paragraph("From somewhere in the wall came the low vibration of an approaching train. Every brass pull trembled at once."),
      new Paragraph("Mara closed the drawer and listened. Then, very carefully, she began to read the names."),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 380 }, children: [new TextRun({ text: "•  •  •", color: "8B5E3C", characterSpacing: 120 })] }),
    ],
  }],
});

// The résumé and pleading-paper presets are NOT generated: those point at
// real-world fixtures (wild3-resume.docx from CareerOneStop, and the
// anonymized pleading-anon.docx) — see PRESETS in src/main.tsx.
const presets = {
  "preset-equations.docx": equations,
  "preset-tables.docx": tables,
  "preset-publication.docx": publication,
  "preset-chapter-book.docx": chapterBook,
};

// The docx library emits <w:rPr> children out of OOXML schema order; sort
// them so Word (and validate-docx.py) accepts the parts. Same fix as
// make-showcase.mjs, applied to every XML part.
const runPropertyOrder = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
  "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
  "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
  "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
  "specVanish", "oMath", "rPrChange",
];
const runPropertyIndex = new Map(runPropertyOrder.map((name, index) => [name, index]));
const sortRunProps = (xml) => xml.replace(
  /<w:rPr>([\s\S]*?)<\/w:rPr>/g,
  (block, contents) => {
    const children = contents.match(/<w:[^>]+\/>/g);
    if (!children || children.join("") !== contents) return block;
    children.sort((a, b) => {
      const aName = a.match(/^<w:([^\s/>]+)/)?.[1] ?? "";
      const bName = b.match(/^<w:([^\s/>]+)/)?.[1] ?? "";
      return (runPropertyIndex.get(aName) ?? 1_000) - (runPropertyIndex.get(bName) ?? 1_000);
    });
    return `<w:rPr>${children.join("")}</w:rPr>`;
  },
);

for (const [name, document] of Object.entries(presets)) {
  const files = unzipSync(new Uint8Array(await Packer.toBuffer(document)));
  for (const part of Object.keys(files)) {
    if (part.endsWith(".xml")) files[part] = strToU8(sortRunProps(strFromU8(files[part])));
  }
  const buffer = zipSync(files);
  writeFileSync(join(outDir, name), buffer);
  console.log("Wrote", name, buffer.length, "bytes");
}
