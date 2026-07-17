#!/usr/bin/env node
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";

const root = join(import.meta.dirname, "..");
const output = join(root, "apps/demo/public/fixtures/insert-editing-parity.docx");

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const documentRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`;

const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

const run = (text, extra = "") => `<w:r>${extra}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const field = (instruction, cached) =>
  `<w:fldSimple w:instr=" ${instruction} "><w:r><w:t xml:space="preserve">${cached}</w:t></w:r></w:fldSimple>`;
const shape = (id, geom, x, y, fill, label, textBox = false) => {
  const cx = 1828800;
  const cy = 914400;
  return `<w:p><w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>
    <wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="${textBox ? "Text Box" : "Shape"} ${id}"/><wp:cNvGraphicFramePr/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>
      ${textBox ? "<a:noFill/>" : `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`}<a:ln w="12700"><a:solidFill><a:srgbClr val="2F5597"/></a:solidFill></a:ln>
    </wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="${textBox ? "202124" : "FFFFFF"}"/></w:rPr><w:t>${label}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>
  </wp:anchor></w:drawing></w:r></w:p>`;
};

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p><w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/></w:pPr>${run("Advanced Insert Editing", "<w:rPr><w:b/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr>")}</w:p>
    <w:p>${run("Bookmark target: ")}<w:bookmarkStart w:id="7" w:name="RevenueTarget"/>${run("Quarterly Revenue", "<w:rPr><w:b/><w:color w:val=\"1F4E79\"/></w:rPr>")}<w:bookmarkEnd w:id="7"/></w:p>
    <w:p>${run("Text cross-reference: ")}${field("REF RevenueTarget \\h \\* MERGEFORMAT", "Quarterly Revenue")}</w:p>
    <w:p>${run("Symbols: Ω   ±   ≤   ≥   ∞   ∑   √   ∫   →   ✓", "<w:rPr><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr>")}</w:p>
    <w:p>${run("Editable equation: ")}
      <m:oMath><m:f><m:num><m:r><m:t>x+1</m:t></m:r></m:num><m:den><m:r><m:t>2y</m:t></m:r></m:den></m:f><m:r><m:t> + </m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>z</m:t></m:r></m:e></m:rad></m:oMath>
    </w:p>
    <w:p>${run("Generic field — total pages: ")}${field("NUMPAGES \\* MERGEFORMAT", "2")}</w:p>
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    <w:p><w:pPr><w:spacing w:after="240"/></w:pPr>${run("Cross-reference continuation", "<w:rPr><w:b/><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr>")}</w:p>
    <w:p>${run("Revenue target is on page ")}${field("PAGEREF RevenueTarget \\h \\* MERGEFORMAT", "1")}${run(".")}</w:p>
    <w:p>${run("The text target still resolves here: ")}${field("REF RevenueTarget \\h \\* MERGEFORMAT", "Quarterly Revenue")}${run(".")}</w:p>
    ${shape(20, "roundRect", 914400, 2743200, "4472C4", "Rounded rectangle")}
    ${shape(21, "ellipse", 3657600, 2743200, "70AD47", "Ellipse")}
    ${shape(22, "diamond", 914400, 4572000, "ED7D31", "Diamond")}
    ${shape(23, "rect", 3657600, 4572000, "FFFFFF", "Editable text box", true)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const files = {
  "[Content_Types].xml": strToU8(contentTypes),
  "_rels/.rels": strToU8(packageRels),
  "word/_rels/document.xml.rels": strToU8(documentRels),
  "word/document.xml": strToU8(document),
  "word/styles.xml": strToU8(styles),
  "word/settings.xml": strToU8(settings),
};

writeFileSync(output, zipSync(files));
console.log(`Wrote ${output}`);
