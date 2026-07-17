#!/usr/bin/env node
import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";
import {
  DocxDocument,
  insertEmbeddedObjectAt,
  insertSmartArtAt,
} from "../../wordinweb/packages/core/dist/index.js";

const fixtureDir = join(dirname(fileURLToPath(import.meta.url)), "../apps/demo/public/fixtures");
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:t xml:space="preserve">Word interoperability validation</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;
const bytes = zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`),
  "word/document.xml": strToU8(documentXml),
});

const poster = Uint8Array.from(Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
  "base64",
));

function generate(name, smartArt, embedded) {
  const doc = DocxDocument.load(bytes);
  const paragraph = doc.sections[0].blocks[0];
  const run = paragraph.children[0];
  const caret = run.content[0].srcT;
  if (!caret) throw new Error("Fixture caret missing");
  if (smartArt) insertSmartArtAt(doc, caret, { layout: "cycle", items: ["Discover", "Design", "Deliver"] });
  if (embedded) {
    insertEmbeddedObjectAt(doc, caret, {
      data: bytes,
      filename: "word-embedded-source.docx",
      poster,
    });
  }
  const output = join(fixtureDir, name);
  writeFileSync(output, doc.save());
  console.log(`Wrote ${output}`);
}

generate("word-interop-smartart-only.docx", true, false);
generate("word-interop-embedded-only.docx", false, true);
generate("word-interop-smartart-embedded.docx", true, true);
generate("word-interop-base-only.docx", false, false);
