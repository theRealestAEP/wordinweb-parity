import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

function malformedCustomPropertiesDocx(): Buffer {
  const files = {
    "[Content_Types].xml": `<?xml version="1.0"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</ct:Types>`,
    "_rels/.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`,
    "word/document.xml": `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Package validity</w:t></w:r></w:p></w:body>
</w:document>`,
    "docProps/custom.xml": `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`,
  };
  return Buffer.from(zipSync(Object.fromEntries(Object.entries(files).map(([name, xml]) => [name, strToU8(xml)]))));
}

test("built-in Download repairs the custom-properties content type exactly once", async ({ page }) => {
  await page.goto("/");
  await page.locator("#docx-upload").setInputFiles({
    name: "malformed-custom-properties.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: malformedCustomPropertiesDocx(),
  });
  await expect(page.getByText("Package validity", { exact: true })).toBeVisible();

  const pending = page.waitForEvent("download");
  await page.getByTitle("Save edited .docx").click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");

  const parts = unzipSync(new Uint8Array(readFileSync(path)));
  expect(parts["docProps/custom.xml"]).toBeDefined();
  const contentTypes = strFromU8(parts["[Content_Types].xml"]);
  expect(contentTypes).toContain('<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">');
  expect(contentTypes.match(/PartName="\/docProps\/custom\.xml"/g) ?? []).toHaveLength(1);
  const overrides = [...contentTypes.matchAll(/<ct:Override\b[^>]*PartName="\/docProps\/custom\.xml"[^>]*>/g)];
  expect(overrides).toHaveLength(1);
  expect(overrides[0][0]).toContain(
    'ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"',
  );
});
