import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  ImageRun,
  Math,
  MathFraction,
  MathIntegral,
  MathRadical,
  MathRun,
  MathSum,
  MathSuperScript,
  PageNumber,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const here = dirname(fileURLToPath(import.meta.url));
const chart = readFileSync(join(here, "assets/showcase-chart.png"));
const output = join(here, "../public/fixtures/showcase.docx");

const borders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "D9E2F3" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "D9E2F3" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "D9E2F3" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "D9E2F3" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D9E2F3" },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D9E2F3" },
};

const formula = (children) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 40, after: 40 },
  children: [new Math({ children })],
});

const headerCell = (text) => new TableCell({
  shading: { type: ShadingType.CLEAR, fill: "17365D" },
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 19 })],
  })],
});

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Calibri", size: 20, color: "243447" } } },
    paragraphStyles: [
      {
        id: "ShowcaseTitle",
        name: "Showcase Title",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Calibri Light", size: 42, bold: true, color: "17365D" },
        paragraph: { spacing: { after: 100 } },
      },
      {
        id: "ShowcaseHeading",
        name: "Showcase Heading",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Calibri", size: 26, bold: true, color: "2F5597" },
        paragraph: { spacing: { before: 130, after: 70 } },
      },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 620, right: 700, bottom: 620, left: 700 } } },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "WordInWeb visual parity demo  •  ", color: "6B7280", size: 17 }), PageNumber.CURRENT],
      })] }),
    },
    children: [
      new Paragraph({ style: "ShowcaseTitle", children: [new TextRun("Scientific document, rendered in JavaScript")] }),
      new Paragraph({
        spacing: { after: 90 },
        children: [
          new TextRun({ text: "Editable DOCX with native Office Math, tables, images, styles, and pagination.  ", size: 22 }),
          new TextRun({ text: "Click anywhere and start typing.", bold: true, color: "2F5597", size: 22 }),
        ],
      }),
      formula([
        new MathRun("E = "),
        new MathFraction({
          numerator: [new MathRun("hc")],
          denominator: [new MathRun("λ")],
        }),
        new MathRun("   ·   "),
        new MathRadical({
          degree: [new MathRun("3")],
          children: [new MathRun("x + 1")],
        }),
        new MathRun("   ·   "),
        new MathSuperScript({ children: [new MathRun("e")], superScript: [new MathRun("iπ")] }),
        new MathRun(" + 1 = 0"),
      ]),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 70, after: 50 },
        children: [new ImageRun({ data: chart, transformation: { width: 650, height: 304 }, type: "png" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 90 },
        children: [new TextRun({ text: "Figure 1. Temperature-dependent diffraction and crystal structure.", italics: true, color: "5B6573", size: 17 })],
      }),
      new Paragraph({ style: "ShowcaseHeading", children: [new TextRun("Equations inside a real Word table")] }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders,
        rows: [
          new TableRow({ children: [headerCell("Model"), headerCell("Native OMML equation"), headerCell("Use") ] }),
          new TableRow({ children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Weighted mean", bold: true })] })] }),
            new TableCell({ children: [formula([
              new MathRun("μ = "),
              new MathFraction({
                numerator: [new MathSum({ children: [new MathRun("wᵢxᵢ")], subScript: [new MathRun("i=1")], superScript: [new MathRun("n")] })],
                denominator: [new MathSum({ children: [new MathRun("wᵢ")], subScript: [new MathRun("i=1")], superScript: [new MathRun("n")] })],
              }),
            ])] }),
            new TableCell({ children: [new Paragraph("Aggregate repeated measurements") ] }),
          ] }),
          new TableRow({ children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Continuous response", bold: true })] })] }),
            new TableCell({ children: [formula([
              new MathIntegral({
                children: [new MathSuperScript({ children: [new MathRun("e")], superScript: [new MathRun("−t²")] }), new MathRun("dt")],
                subScript: [new MathRun("0")],
                superScript: [new MathRun("∞")],
              }),
            ])] }),
            new TableCell({ children: [new Paragraph("Model a smooth distribution") ] }),
          ] }),
        ],
      }),
      new Paragraph({
        spacing: { before: 120 },
        children: [new TextRun({
          text: "Everything on this page remains editable and downloads as a DOCX—no PDF conversion and no server-side rendering.",
          color: "40566F",
          size: 19,
        })],
      }),
    ],
  }],
});

mkdirSync(dirname(output), { recursive: true });
const files = unzipSync(new Uint8Array(await Packer.toBuffer(doc)));
const runPropertyOrder = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
  "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
  "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
  "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
  "specVanish", "oMath", "rPrChange",
];
const runPropertyIndex = new Map(runPropertyOrder.map((name, index) => [name, index]));
const styles = strFromU8(files["word/styles.xml"]).replace(
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
files["word/styles.xml"] = strToU8(styles);
writeFileSync(output, zipSync(files));
