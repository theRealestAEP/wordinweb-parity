// Generates test .docx fixtures into apps/demo/public/fixtures using the
// `docx` package (dev-only dependency; the viewer itself never uses it).
import {
  AlignmentType,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  TextWrappingType,
  VerticalPositionRelativeFrom,
  BorderStyle,
  ImageRun,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  Document,
  Footer,
  FootnoteReferenceRun,
  Header,
  HeadingLevel,
  LevelFormat,
  PageBreak,
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
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const outDir = join(dirname(fileURLToPath(import.meta.url)), "../public/fixtures");
mkdirSync(outDir, { recursive: true });

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22 } }, // 11pt
    },
  },
  comments: {
    children: [
      {
        id: 0,
        author: "Ada Reviewer",
        initials: "AR",
        date: new Date("2026-06-01T10:00:00Z"),
        children: [new Paragraph({ children: [new TextRun("Should this be brand blue instead of red?")] })],
      },
      {
        id: 1,
        author: "Bob Editor",
        initials: "BE",
        date: new Date("2026-06-02T15:30:00Z"),
        children: [new Paragraph({ children: [new TextRun("Numbering restarts here — double-check the list level.")] })],
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "num-list",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
          {
            level: 1,
            format: LevelFormat.LOWER_LETTER,
            text: "%2)",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: "DocxInWeb Fidelity Sample", italics: true, size: 18 })],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 18 }),
              ],
            }),
          ],
        }),
      },
      children: [
        new Paragraph({ text: "DocxInWeb Rendering Test", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          children: [
            new TextRun("This document exercises the fidelity-critical features: "),
            new TextRun({ text: "bold", bold: true }),
            new TextRun(", "),
            new TextRun({ text: "italic", italics: true }),
            new TextRun(", "),
            new TextRun({ text: "underline", underline: {} }),
            new TextRun(", "),
            new CommentRangeStart(0),
            new TextRun({ text: "colored text", color: "C00000" }),
            new CommentRangeEnd(0),
            new TextRun({ children: [new CommentReference(0)] }),
            new TextRun(", "),
            new TextRun({ text: "highlight", highlight: "yellow" }),
            new TextRun(", and "),
            new TextRun({ text: "superscript", superScript: true }),
            new TextRun("."),
          ],
        }),
        new Paragraph({ text: "Justified paragraph", heading: HeadingLevel.HEADING_2 }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          children: [
            new TextRun(
              "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.",
            ),
          ],
        }),
        new Paragraph({ text: "Lists", heading: HeadingLevel.HEADING_2 }),
        new Paragraph({ text: "First numbered item", numbering: { reference: "num-list", level: 0 } }),
        new Paragraph({
          numbering: { reference: "num-list", level: 0 },
          children: [
            new CommentRangeStart(1),
            new TextRun("Second numbered item"),
            new CommentRangeEnd(1),
            new TextRun({ children: [new CommentReference(1)] }),
          ],
        }),
        new Paragraph({ text: "Nested letter item", numbering: { reference: "num-list", level: 1 } }),
        new Paragraph({ text: "Another nested item", numbering: { reference: "num-list", level: 1 } }),
        new Paragraph({ text: "Back to top level", numbering: { reference: "num-list", level: 0 } }),
        new Paragraph({ text: "Bullet one", bullet: { level: 0 } }),
        new Paragraph({ text: "Bullet two", bullet: { level: 0 } }),
        new Paragraph({ text: "Divider line below (paragraph bottom border):" }),
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1, color: "auto" } },
          text: "",
        }),
        new Paragraph({ text: "Table", heading: HeadingLevel.HEADING_2 }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              tableHeader: true,
              children: ["Feature", "Status", "Notes"].map(
                (t) =>
                  new TableCell({
                    shading: { type: ShadingType.CLEAR, fill: "D9E2F3" },
                    children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })],
                  }),
              ),
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("Pagination")] }),
                new TableCell({ children: [new Paragraph("Working")] }),
                new TableCell({ children: [new Paragraph("Real page boxes with measured line breaking")] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("Page numbers")] }),
                new TableCell({ children: [new Paragraph("Working")] }),
                new TableCell({ children: [new Paragraph("PAGE / NUMPAGES fields resolved at layout time")] }),
              ],
            }),
          ],
        }),
        new Paragraph({ children: [new PageBreak()], text: "" }),
        new Paragraph({ text: "Page 2", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          text: "This paragraph starts page two after an explicit page break. The footer below should read “Page 2 of N”.",
        }),
        ...Array.from({ length: 60 }, (_, i) =>
          new Paragraph({
            text: `Filler paragraph ${i + 1} — long enough content to force natural pagination across multiple pages so that widow control, page fill and footers can be verified visually.`,
          }),
        ),
      ],
    },
  ],
});

const buf = await Packer.toBuffer(doc);
writeFileSync(join(outDir, "sample.docx"), buf);
console.log("Wrote", join(outDir, "sample.docx"), buf.length, "bytes");

// Pleading-style fixture: every paragraph on a fixed 24pt grid via exact
// line spacing (like California pleading paper). Exercises image handling
// in fixed-height lines — images that can't fit must float with wrap.
const exactDoc = new Document({
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } }, // 12pt
    },
  },
  sections: [
    {
      children: Array.from(
        { length: 40 },
        (_, i) =>
          new Paragraph({
            spacing: { line: 480, lineRule: "exact", before: 0, after: 0 },
            children: [
              new TextRun(
                `Exact line ${i + 1}: the quick brown fox jumps over the lazy dog near the riverbank at dawn.`,
              ),
            ],
          }),
      ),
    },
  ],
});

const exactBuf = await Packer.toBuffer(exactDoc);
writeFileSync(join(outDir, "exact.docx"), exactBuf);
console.log("Wrote", join(outDir, "exact.docx"), exactBuf.length, "bytes");

// Footnote fixture: refs on page 1 and (after filler) page 2, so the
// same-page binding and bottom-of-body placement can be checked visually.
const fnDoc = new Document({
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
  },
  footnotes: {
    1: {
      children: [
        new Paragraph({
          children: [
            new TextRun(
              "This is the first footnote. It has enough text to wrap onto a second line so the footnote area height is visible in the demo.",
            ),
          ],
        }),
      ],
    },
    2: {
      children: [
        new Paragraph({
          children: [new TextRun("The second footnote belongs to the second page, right where its reference is.")],
        }),
      ],
    },
  },
  sections: [
    {
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        new Paragraph({
          children: [
            new TextRun("The first footnote reference appears in the very first paragraph of the body"),
            new FootnoteReferenceRun(1),
            new TextRun(" and the sentence continues after the mark."),
          ],
        }),
        ...Array.from(
          { length: 46 },
          (_, i) =>
            new Paragraph({
              text: `This is body filler paragraph number ${i + 1}, present to push the second footnote reference onto the second page of the document.`,
            }),
        ),
        new Paragraph({
          children: [
            new TextRun("A second reference lands on the following page"),
            new FootnoteReferenceRun(2),
            new TextRun(" and its note must appear at the bottom of that same page."),
          ],
        }),
        new Paragraph({ text: "Final body paragraph." }),
      ],
    },
  ],
});

const fnBuf = await Packer.toBuffer(fnDoc);
writeFileSync(join(outDir, "footnotes.docx"), fnBuf);
console.log("Wrote", join(outDir, "footnotes.docx"), fnBuf.length, "bytes");

// ---------------------------------------------------------------------------
// Focused parity fixtures: one feature area per file, so Word-vs-web diffs
// (npm run parity) attribute regressions to a specific subsystem.
// Export references with: scripts/word-parity.sh apps/demo/public/fixtures/parity-<x>.docx

const CAL = { styles: { default: { document: { run: { font: "Calibri", size: 22 } } } } };
const LOREM =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.";

// 40x24 solid red PNG for picture fixtures.
const RED_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAACgAAAAYCAIAAAAH5iiXAAAAKklEQVR42u3NQREAAAQAMCTRP4VoSrjz2Qospzs+VDwRi8VisVgsFovvLXLsATBmdMUuAAAAAElFTkSuQmCC",
  "base64",
);

const parityDocs = {
  "parity-text": new Document({
    ...CAL,
    sections: [
      {
        children: [
          new Paragraph({ text: "Plain text parity", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun(LOREM)] }),
          new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun(LOREM)] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("Centered single line of text")] }),
          new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun("Right-aligned single line")] }),
          new Paragraph({
            children: [
              new TextRun("Mixed: "),
              new TextRun({ text: "bold", bold: true }),
              new TextRun({ text: " italic", italics: true }),
              new TextRun({ text: " 16pt", size: 32 }),
              new TextRun({ text: " colored", color: "C00000" }),
              new TextRun({ text: " highlighted", highlight: "yellow" }),
              new TextRun("."),
            ],
          }),
        ],
      },
    ],
  }),

  "parity-lists": new Document({
    ...CAL,
    numbering: {
      config: [
        {
          reference: "pl",
          levels: [
            { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
            { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
          ],
        },
      ],
    },
    sections: [
      {
        children: [
          new Paragraph({ text: "Lists parity", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: "First numbered item", numbering: { reference: "pl", level: 0 } }),
          new Paragraph({ text: "Second numbered item", numbering: { reference: "pl", level: 0 } }),
          new Paragraph({ text: "Nested letter item", numbering: { reference: "pl", level: 1 } }),
          new Paragraph({ text: "Back to top level", numbering: { reference: "pl", level: 0 } }),
          new Paragraph({ text: "Bullet one", bullet: { level: 0 } }),
          new Paragraph({ text: "Bullet two", bullet: { level: 0 } }),
          new Paragraph({ text: "Nested bullet", bullet: { level: 1 } }),
        ],
      },
    ],
  }),

  "parity-tables": new Document({
    ...CAL,
    sections: [
      {
        children: [
          new Paragraph({ text: "Tables parity", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: "Autofit (100% width, content-sized columns):" }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                tableHeader: true,
                children: ["Key", "Status", "Description of the item"].map(
                  (t) =>
                    new TableCell({
                      shading: { type: ShadingType.CLEAR, fill: "D9E2F3" },
                      children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })],
                    }),
                ),
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("A")] }),
                  new TableCell({ children: [new Paragraph("ok")] }),
                  new TableCell({ children: [new Paragraph("A much longer description cell that should dominate the width")] }),
                ],
              }),
            ],
          }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: "Fixed columns (2in / 4in):" }),
          new Table({
            columnWidths: [2880, 5760],
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph("Left 2in")] }),
                  new TableCell({ children: [new Paragraph("Right 4in with some wrapping content to check the row height")] }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ shading: { type: ShadingType.CLEAR, fill: "FFF2CC" }, children: [new Paragraph("Shaded")] }),
                  new TableCell({ children: [new Paragraph("Plain")] }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  }),

  "parity-comments": new Document({
    ...CAL,
    comments: {
      children: [
        { id: 0, author: "Ada Reviewer", initials: "AR", date: new Date("2026-06-01T10:00:00Z"), children: [new Paragraph({ children: [new TextRun("First comment body.")] })] },
        { id: 1, author: "Bob Editor", initials: "BE", date: new Date("2026-06-02T11:00:00Z"), children: [new Paragraph({ children: [new TextRun("Second comment, on a different paragraph.")] })] },
      ],
    },
    sections: [
      {
        children: [
          new Paragraph({ text: "Comments parity", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun("A paragraph with "),
              new CommentRangeStart(0),
              new TextRun("a commented range"),
              new CommentRangeEnd(0),
              new TextRun({ children: [new CommentReference(0)] }),
              new TextRun(" in the middle."),
            ],
          }),
          new Paragraph({ children: [new TextRun(LOREM)] }),
          new Paragraph({
            children: [
              new CommentRangeStart(1),
              new TextRun("This whole sentence is commented."),
              new CommentRangeEnd(1),
              new TextRun({ children: [new CommentReference(1)] }),
            ],
          }),
        ],
      },
    ],
  }),

  "parity-headerfooter": new Document({
    ...CAL,
    sections: [
      {
        headers: {
          default: new Header({
            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Header/Footer Parity", italics: true, size: 18 })] })],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 18 })],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({ text: "Header & footer parity", heading: HeadingLevel.HEADING_1 }),
          ...Array.from({ length: 70 }, (_, i) => new Paragraph({ text: `Filler paragraph ${i + 1} pushing content across several pages so headers and footers repeat.` })),
        ],
      },
    ],
  }),

  "parity-dividers": new Document({
    ...CAL,
    sections: [
      {
        children: [
          new Paragraph({ text: "Dividers parity", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: "Thin single divider below:" }),
          new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1, color: "auto" } }, text: "" }),
          new Paragraph({ text: "Thick divider below:" }),
          new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 24, space: 1, color: "999999" } }, text: "" }),
          new Paragraph({ text: "Double divider below:" }),
          new Paragraph({ border: { bottom: { style: BorderStyle.DOUBLE, size: 6, space: 1, color: "auto" } }, text: "" }),
          new Paragraph({ text: "Dashed divider below:" }),
          new Paragraph({ border: { bottom: { style: BorderStyle.DASHED, size: 6, space: 1, color: "2E74B5" } }, text: "" }),
          new Paragraph({ text: "End." }),
        ],
      },
    ],
  }),

  "parity-pictures": new Document({
    ...CAL,
    sections: [
      {
        children: [
          new Paragraph({ text: "Pictures parity", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun("Inline image between text "),
              new ImageRun({ data: RED_PNG, transformation: { width: 60, height: 36 }, type: "png" }),
              new TextRun(" and the sentence continues."),
            ],
          }),
          new Paragraph({ children: [new TextRun(LOREM)] }),
          new Paragraph({
            children: [new ImageRun({ data: RED_PNG, transformation: { width: 160, height: 96 }, type: "png" })],
          }),
          new Paragraph({ children: [new TextRun("Text after a block image.")] }),
        ],
      },
    ],
  }),

  "parity-firstpage": new Document({
    ...CAL,
    sections: [
      {
        properties: { titlePage: true },
        headers: {
          first: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "FIRST PAGE HEADER", bold: true, size: 20 })] })] }),
          default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Running header", italics: true, size: 18 })] })] }),
        },
        footers: {
          first: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "First footer", size: 18 })] })] }),
          default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", PageNumber.CURRENT], size: 18 })] })] }),
        },
        children: [
          new Paragraph({ text: "First-page header fixture", heading: HeadingLevel.HEADING_1 }),
          ...Array.from({ length: 60 }, (_, i) =>
            new Paragraph({ children: [new TextRun(`Body filler paragraph ${i + 1} pushing content onto a second page so both header variants show.`)] }),
          ),
        ],
      },
    ],
  }),

  "parity-wrapmodes": new Document({
    ...CAL,
    sections: [
      {
        children: [
          new Paragraph({ text: "Wrap modes fixture", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new ImageRun({
                data: RED_PNG,
                transformation: { width: 120, height: 72 },
                type: "png",
                floating: {
                  horizontalPosition: { relative: HorizontalPositionRelativeFrom.MARGIN, align: HorizontalPositionAlign.LEFT },
                  verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
                  wrap: { type: TextWrappingType.SQUARE },
                },
              }),
              new TextRun(LOREM + " " + LOREM),
            ],
          }),
          new Paragraph({
            children: [
              new ImageRun({
                data: RED_PNG,
                transformation: { width: 140, height: 84 },
                type: "png",
                floating: {
                  horizontalPosition: { relative: HorizontalPositionRelativeFrom.MARGIN, align: HorizontalPositionAlign.CENTER },
                  verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
                  wrap: { type: TextWrappingType.NONE },
                  behindDocument: true,
                },
              }),
              new TextRun("Text over a behind-document image: " + LOREM),
            ],
          }),
        ],
      },
    ],
  }),

  benchmark: new Document({
    ...CAL,
    numbering: {
      config: [
        {
          reference: "bench-num",
          levels: [
            { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
            { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
          ],
        },
        {
          reference: "bench-bullets",
          levels: [
            { level: 0, format: LevelFormat.BULLET, text: "\u25cf", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          ],
        },
      ],
    },
    sections: [
      {
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: "DocxInWeb Benchmark", italics: true, size: 18 })],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 18 })],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({ text: "The Kitchen Sink Benchmark", heading: HeadingLevel.TITLE }),
          new Paragraph({
            children: [new TextRun({ text: "Every feature on a few pages: a general parity yardstick.", italics: true, color: "5f6368" })],
          }),
          new Paragraph({ text: "Typography", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun("Plain, "),
              new TextRun({ text: "bold, ", bold: true }),
              new TextRun({ text: "italic, ", italics: true }),
              new TextRun({ text: "underlined, ", underline: {} }),
              new TextRun({ text: "struck, ", strike: true }),
              new TextRun({ text: "highlighted, ", highlight: "yellow" }),
              new TextRun({ text: "colored, ", color: "C00000" }),
              new TextRun({ text: "x", }),
              new TextRun({ text: "2", superScript: true }),
              new TextRun({ text: " and H" }),
              new TextRun({ text: "2", subScript: true }),
              new TextRun({ text: "O." }),
            ],
          }),
          new Paragraph({ children: [new TextRun({ text: "Cambria at 14pt for a serif line of text.", font: "Cambria", size: 28 })] }),
          new Paragraph({ children: [new TextRun({ text: "Arial at 10pt, the workhorse sans.", font: "Arial", size: 20 })] }),
          new Paragraph({ children: [new TextRun({ text: "Times New Roman at 12pt, the classic.", font: "Times New Roman", size: 24 })] }),
          new Paragraph({ children: [new TextRun({ text: "Georgia at 11pt brings wide digits 0123456789.", font: "Georgia", size: 22 })] }),
          new Paragraph({ children: [new TextRun({ text: "Courier New at 10pt is monospaced.", font: "Courier New", size: 20 })] }),
          new Paragraph({ text: "Alignment and flow", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ alignment: AlignmentType.BOTH, children: [new TextRun(LOREM + " " + LOREM)] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("A centered line sits in the middle.")] }),
          new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun("And this one hugs the right margin.")] }),
          new Paragraph({ text: "Lists", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "First numbered item", numbering: { reference: "bench-num", level: 0 } }),
          new Paragraph({ text: "Second numbered item", numbering: { reference: "bench-num", level: 0 } }),
          new Paragraph({ text: "Nested letter item", numbering: { reference: "bench-num", level: 1 } }),
          new Paragraph({ text: "Bullet one", numbering: { reference: "bench-bullets", level: 0 } }),
          new Paragraph({ text: "Bullet two with a somewhat longer text body to wrap around", numbering: { reference: "bench-bullets", level: 0 } }),
          new Paragraph({ text: "Tables", heading: HeadingLevel.HEADING_1 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                tableHeader: true,
                children: ["Feature", "Status", "Notes"].map(
                  (t) =>
                    new TableCell({
                      shading: { type: ShadingType.CLEAR, fill: "D9E2F3" },
                      children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })],
                    }),
                ),
              }),
              new TableRow({
                children: ["Pagination", "done", "Widow control and fills included"].map(
                  (t) => new TableCell({ children: [new Paragraph(t)] }),
                ),
              }),
              new TableRow({
                children: ["Justify", "done", "Pack-vs-break rule measured from Word"].map(
                  (t) => new TableCell({ children: [new Paragraph(t)] }),
                ),
              }),
            ],
          }),
          new Paragraph({ text: "" }),
          new Table({
            columnWidths: [2880, 5760],
            rows: [
              new TableRow({
                children: [
                  new TableCell({ width: { size: 2880, type: WidthType.DXA }, children: [new Paragraph("Fixed 2in")] }),
                  new TableCell({
                    width: { size: 5760, type: WidthType.DXA },
                    children: [new Paragraph("Fixed 4in column with wrapping content to give the row some height")],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ width: { size: 2880, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: "FFF2CC" }, children: [new Paragraph("Shaded")] }),
                  new TableCell({ width: { size: 5760, type: WidthType.DXA }, children: [new Paragraph("Plain")] }),
                ],
              }),
            ],
          }),
          new Paragraph({ text: "Pictures", heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun("An inline image "),
              new ImageRun({ data: RED_PNG, transformation: { width: 60, height: 36 }, type: "png" }),
              new TextRun(" interrupts this sentence, which then carries on to the end of the line."),
            ],
          }),
          new Paragraph({
            children: [
              new ImageRun({
                data: RED_PNG,
                transformation: { width: 140, height: 84 },
                type: "png",
                floating: {
                  horizontalPosition: { relative: HorizontalPositionRelativeFrom.MARGIN, align: HorizontalPositionAlign.RIGHT },
                  verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
                  wrap: { type: TextWrappingType.SQUARE },
                },
              }),
              new TextRun(LOREM + " " + LOREM),
            ],
          }),
          new Paragraph({ children: [new PageBreak()], text: "" }),
          new Paragraph({ text: "Deep pages", heading: HeadingLevel.HEADING_1 }),
          ...Array.from({ length: 30 }, (_, i) =>
            new Paragraph({
              alignment: i % 3 === 0 ? AlignmentType.BOTH : AlignmentType.LEFT,
              children: [
                new TextRun(
                  `Benchmark filler paragraph ${i + 1} - long enough content to force natural pagination across several pages so headers, footers and page numbers can be checked everywhere. `,
                ),
                new TextRun(i % 5 === 0 ? LOREM : "The quick brown fox jumps over the lazy dog near the riverbank at dawn."),
              ],
            }),
          ),
          new Paragraph({ text: "A closing heading on the last page", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ children: [new TextRun("The end of the benchmark document.")] }),
        ],
      },
    ],
  }),
};

for (const [name, d] of Object.entries(parityDocs)) {
  const buf = await Packer.toBuffer(d);
  writeFileSync(join(outDir, name + ".docx"), buf);
  console.log("Wrote", join(outDir, name + ".docx"), buf.length, "bytes");
}
