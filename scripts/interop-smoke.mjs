#!/usr/bin/env node

import { createSign } from "node:crypto";
import { access, copyFile, mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";
import { DocxDocument } from "wordinweb";

const FIXTURE_DIR = resolve("apps/demo/public/fixtures");
const REPORT_DIR = resolve("apps/demo/public/interop");
const REPORT_MANIFEST = join(REPORT_DIR, "results.json");
const FIXTURES = [
  "parity-tables.docx",
  "parity2-nestedtables.docx",
  "parity2-tabs.docx",
  "probe3-table-exotics.docx",
  "preset-tables.docx",
];
const includeGoogle = process.argv.includes("--google");
const keepArtifacts = process.argv.includes("--keep");

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: options.encoding ?? "utf8",
    maxBuffer: 32 * 1024 * 1024,
    ...options,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} exited ${result.status}\n${result.stdout ?? ""}\n${result.stderr ?? ""}`);
  }
  return result;
}

function textOfParagraph(paragraph) {
  let text = "";
  for (const child of paragraph.children ?? []) {
    const runs = child.type === "run" ? [child] : child.runs ?? [];
    for (const oneRun of runs) {
      for (const content of oneRun.content ?? []) {
        if (content.kind === "text") text += content.text;
      }
    }
  }
  return text;
}

function docxStats(bytes) {
  const document = DocxDocument.load(bytes);
  const stats = { tables: 0, rows: 0, cells: 0, tabStops: 0, tabs: 0, text: "" };
  const walkBlocks = (blocks) => {
    for (const block of blocks ?? []) {
      if (block.type === "paragraph") {
        stats.text += `${textOfParagraph(block)}\n`;
        stats.tabStops += block.props?.tabs?.length ?? 0;
        for (const child of block.children ?? []) {
          const runs = child.type === "run" ? [child] : child.runs ?? [];
          for (const oneRun of runs) {
            stats.tabs += (oneRun.content ?? []).filter((content) => content.kind === "tab").length;
          }
        }
      } else if (block.type === "table") {
        stats.tables += 1;
        stats.rows += block.rows.length;
        for (const row of block.rows) {
          stats.cells += row.cells.length;
          for (const cell of row.cells) walkBlocks(cell.blocks);
        }
      }
    }
  };
  for (const section of document.sections) walkBlocks(section.blocks);
  stats.text = stats.text.replace(/\s+/g, " ").trim();
  return stats;
}

function structureSummary(stats) {
  return `${stats.tables} tables, ${stats.rows} rows, ${stats.cells} cells, ${stats.tabStops} tab stops, ${stats.tabs} tabs, ${stats.text.length} text characters`;
}

function assertStructure(reference, converted, label, minimumRatio = 0.75) {
  if (reference.tables > 0) {
    assert(converted.tables >= Math.ceil(reference.tables * minimumRatio),
      `${label}: table count collapsed (${converted.tables} from ${reference.tables})`);
    assert(converted.rows >= Math.ceil(reference.rows * minimumRatio),
      `${label}: row count collapsed (${converted.rows} from ${reference.rows})`);
    assert(converted.cells >= Math.ceil(reference.cells * minimumRatio),
      `${label}: cell count collapsed (${converted.cells} from ${reference.cells})`);
  }
  if (reference.tabStops > 0) {
    assert(converted.tabStops >= Math.ceil(reference.tabStops * minimumRatio),
      `${label}: tab stops collapsed (${converted.tabStops} from ${reference.tabStops})`);
    assert(converted.tabs >= Math.ceil(reference.tabs * minimumRatio),
      `${label}: tab characters collapsed (${converted.tabs} from ${reference.tabs})`);
  }
  assert(converted.text.length >= Math.ceil(reference.text.length * 0.65),
    `${label}: document text collapsed (${converted.text.length} from ${reference.text.length} characters)`);

  const referenceTokens = [...new Set(reference.text.match(/[\p{L}\p{N}]{4,}/gu) ?? [])].slice(0, 12);
  const retained = referenceTokens.filter((token) => converted.text.includes(token));
  assert(retained.length >= Math.ceil(referenceTokens.length * 0.5),
    `${label}: representative table text was lost (${retained.length}/${referenceTokens.length} tokens retained)`);
}

function pdfInfo(path) {
  const output = run(process.env.PDFINFO_BIN ?? "pdfinfo", [path]).stdout;
  const pages = Number(/^Pages:\s+(\d+)/m.exec(output)?.[1] ?? 0);
  assert(pages > 0, `${basename(path)}: PDF has no pages`);
  return { pages };
}

function pdfText(path) {
  return run(process.env.PDFTOTEXT_BIN ?? "pdftotext", ["-layout", path, "-"]).stdout
    .replace(/\s+/g, " ")
    .trim();
}

async function ppmInk(path, workDir) {
  const outputBase = join(workDir, `${basename(path, ".pdf")}-first-page`);
  run(process.env.PDFTOPPM_BIN ?? "pdftoppm", ["-f", "1", "-singlefile", "-r", "72", path, outputBase]);
  const bytes = await readFile(`${outputBase}.ppm`);
  let offset = 0;
  const token = () => {
    while (offset < bytes.length) {
      if (bytes[offset] === 35) {
        while (offset < bytes.length && bytes[offset] !== 10) offset += 1;
      } else if (bytes[offset] <= 32) offset += 1;
      else break;
    }
    const start = offset;
    while (offset < bytes.length && bytes[offset] > 32) offset += 1;
    return bytes.subarray(start, offset).toString("ascii");
  };
  assert(token() === "P6", `${basename(path)}: expected a binary PPM render`);
  const width = Number(token());
  const height = Number(token());
  assert(Number(token()) === 255, `${basename(path)}: unsupported PPM color depth`);
  while (bytes[offset] <= 32) offset += 1;

  let ink = 0;
  let sampled = 0;
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;
  const pixels = bytes.subarray(offset);
  for (let pixel = 0; pixel + 2 < pixels.length; pixel += 12) {
    const index = pixel / 3;
    const x = index % width;
    const y = Math.floor(index / width);
    sampled += 1;
    if (pixels[pixel] < 245 || pixels[pixel + 1] < 245 || pixels[pixel + 2] < 245) {
      ink += 1;
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  }
  return {
    width,
    height,
    inkRatio: ink / sampled,
    inkWidth: maxX >= minX ? maxX - minX + 1 : 0,
    inkHeight: maxY >= minY ? maxY - minY + 1 : 0,
  };
}

function assertPdf(reference, candidate, label) {
  assert(candidate.pages >= Math.max(1, Math.ceil(reference.pages * 0.75)),
    `${label}: page count collapsed (${candidate.pages} from ${reference.pages})`);
  assert(candidate.pages <= reference.pages + Math.floor(reference.pages * 0.25),
    `${label}: page count exploded (${candidate.pages} from ${reference.pages})`);
  assert(candidate.text.length >= Math.ceil(reference.text.length * 0.4),
    `${label}: exported PDF text collapsed (${candidate.text.length} from ${reference.text.length} characters)`);
  assert(candidate.ink.inkRatio >= Math.max(0.001, reference.ink.inkRatio * 0.2),
    `${label}: first page is nearly blank (${candidate.ink.inkRatio.toFixed(4)} ink ratio)`);
  assert(candidate.ink.inkWidth >= candidate.ink.width * 0.1,
    `${label}: first-page content collapsed horizontally`);
  assert(candidate.ink.inkHeight >= candidate.ink.height * 0.05,
    `${label}: first-page content collapsed vertically`);
}

async function inspectPdf(path, workDir) {
  return {
    ...pdfInfo(path),
    text: pdfText(path),
    ink: await ppmInk(path, workDir),
  };
}

async function renderPdfPreviews(path, fixture, engine, pages) {
  await mkdir(REPORT_DIR, { recursive: true });
  const stem = `${basename(fixture, ".docx")}-${engine}`;
  run(process.env.PDFTOPPM_BIN ?? "pdftoppm", [
    "-png",
    "-r",
    "96",
    path,
    join(REPORT_DIR, stem),
  ]);
  return Array.from({ length: pages }, (_, index) => `interop/${stem}-${index + 1}.png`);
}

async function copyWordInWebPreviews(fixture, pages) {
  await mkdir(REPORT_DIR, { recursive: true });
  const stem = basename(fixture, ".docx");
  const previews = [];
  for (let page = 1; page <= pages; page += 1) {
    const name = `${stem}-wordinweb-${page}.png`;
    const source = resolve("parity", "out", `${stem}-p${page}.png`);
    const [compositeWidth, height] = run("magick", ["identify", "-format", "%w %h", source]).stdout
      .trim()
      .split(/\s+/)
      .map(Number);
    const gap = 12;
    const pageWidth = (compositeWidth - gap * 2) / 3;
    assert(Number.isInteger(pageWidth), `${basename(source)}: unexpected parity image width`);
    run("magick", [
      source,
      "-crop",
      `${pageWidth}x${height}+${pageWidth + gap}+0`,
      "+repage",
      join(REPORT_DIR, name),
    ]);
    previews.push(`interop/${name}`);
  }
  return previews;
}

async function writeReportManifest(results) {
  let previous = { results: [] };
  try {
    previous = JSON.parse(await readFile(REPORT_MANIFEST, "utf8"));
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
  const previousByFixture = new Map(previous.results.map((result) => [result.fixture, result]));
  const merged = results.map((result) => ({
    ...previousByFixture.get(result.fixture),
    ...result,
    ...(!includeGoogle && previousByFixture.get(result.fixture)?.googleDocs
      ? { googleDocs: previousByFixture.get(result.fixture).googleDocs }
      : {}),
  }));
  await mkdir(REPORT_DIR, { recursive: true });
  await writeFile(REPORT_MANIFEST, `${JSON.stringify({ generatedAt: new Date().toISOString(), results: merged }, null, 2)}\n`);
}

async function libreOfficeConvert(input, format, outputDir, profileDir) {
  await mkdir(outputDir, { recursive: true });
  const soffice = process.env.SOFFICE_BIN ?? "soffice";
  run(soffice, [
    `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
    "--headless",
    "--convert-to",
    format,
    "--outdir",
    outputDir,
    input,
  ]);
  const output = join(outputDir, `${basename(input, ".docx")}.${format}`);
  await readFile(output);
  return output;
}

function base64Url(value) {
  return Buffer.from(value).toString("base64url");
}

async function googleAccessToken() {
  if (process.env.GOOGLE_INTEROP_ACCESS_TOKEN) return process.env.GOOGLE_INTEROP_ACCESS_TOKEN;
  const credentialSource = process.env.GOOGLE_INTEROP_SERVICE_ACCOUNT_JSON;
  assert(credentialSource, "Set GOOGLE_INTEROP_ACCESS_TOKEN or GOOGLE_INTEROP_SERVICE_ACCOUNT_JSON for --google");
  const raw = credentialSource.trim().startsWith("{")
    ? credentialSource
    : await readFile(resolve(credentialSource), "utf8");
  const credentials = JSON.parse(raw);
  const now = Math.floor(Date.now() / 1000);
  const claim = `${base64Url(JSON.stringify({ alg: "RS256", typ: "JWT" }))}.${base64Url(JSON.stringify({
    iss: credentials.client_email,
    scope: "https://www.googleapis.com/auth/drive",
    aud: credentials.token_uri ?? "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  }))}`;
  const signer = createSign("RSA-SHA256");
  signer.update(claim);
  signer.end();
  const assertion = `${claim}.${signer.sign(credentials.private_key).toString("base64url")}`;
  const response = await fetch(credentials.token_uri ?? "https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion,
    }),
  });
  assert(response.ok, `Google OAuth failed (${response.status}): ${await response.text()}`);
  return (await response.json()).access_token;
}

async function googleRequest(token, url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: { authorization: `Bearer ${token}`, ...options.headers },
  });
  assert(response.ok, `Google Drive request failed (${response.status}): ${await response.text()}`);
  return response;
}

async function googleImport(token, path, title) {
  const boundary = `wordinweb-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const metadata = {
    name: title,
    mimeType: "application/vnd.google-apps.document",
    ...(process.env.GOOGLE_INTEROP_FOLDER_ID ? { parents: [process.env.GOOGLE_INTEROP_FOLDER_ID] } : {}),
  };
  const source = await readFile(path);
  const body = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`),
    Buffer.from(`--${boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n`),
    source,
    Buffer.from(`\r\n--${boundary}--\r\n`),
  ]);
  const response = await googleRequest(token,
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,mimeType",
    { method: "POST", headers: { "content-type": `multipart/related; boundary=${boundary}` }, body });
  return response.json();
}

async function googleExport(token, id, mimeType) {
  const response = await googleRequest(token,
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}/export?mimeType=${encodeURIComponent(mimeType)}`);
  return new Uint8Array(await response.arrayBuffer());
}

async function googleDelete(token, id) {
  await googleRequest(token,
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?supportsAllDrives=true`,
    { method: "DELETE" });
}

async function main() {
  run(process.env.SOFFICE_BIN ?? "soffice", ["--version"]);
  run(process.env.PDFINFO_BIN ?? "pdfinfo", ["-v"]);
  const workDir = await mkdtemp(join(tmpdir(), "wordinweb-interop-"));
  const profileDir = join(workDir, "libreoffice-profile");
  const googleToken = includeGoogle ? await googleAccessToken() : null;
  const failures = [];
  const reportResults = [];

  console.log(`Cross-editor layout smoke suite: ${FIXTURES.length} fixtures`);
  console.log(`Artifacts: ${workDir}`);
  try {
    for (const fixture of FIXTURES) {
      const reportResult = { fixture };
      let activeEngine = "libreOffice";
      try {
        const sourcePath = join(FIXTURE_DIR, fixture);
        const sourceBytes = new Uint8Array(await readFile(sourcePath));
        const sourceStats = docxStats(sourceBytes);
        assert(sourceStats.tables > 0 || sourceStats.tabStops > 0,
          `${fixture}: fixture must contain a table or explicit tab stop`);

        const wordReferencePath = resolve("parity", `${basename(fixture, ".docx")}-word.pdf`);
        let wordReference = null;
        try {
          await access(wordReferencePath);
          wordReference = await inspectPdf(wordReferencePath, workDir);
          reportResult.reference = {
            label: "WordInWeb",
            pages: wordReference.pages,
            previews: await copyWordInWebPreviews(fixture, wordReference.pages),
          };
        } catch (error) {
          if (error?.code !== "ENOENT") throw error;
        }

        const candidatePath = join(workDir, `${basename(fixture, ".docx")}-wordinweb.docx`);
        const candidateBytes = DocxDocument.load(sourceBytes).save();
        await writeFile(candidatePath, candidateBytes);
        const candidateStats = docxStats(candidateBytes);
        assertStructure(sourceStats, candidateStats, `${fixture} WordInWeb save`, 1);

        const sourcePdfPath = await libreOfficeConvert(sourcePath, "pdf", join(workDir, "lo-source-pdf"), profileDir);
        const candidatePdfPath = await libreOfficeConvert(candidatePath, "pdf", join(workDir, "lo-candidate-pdf"), profileDir);
        const roundTripPath = await libreOfficeConvert(candidatePath, "docx", join(workDir, "lo-roundtrip-docx"), profileDir);
        const sourcePdf = await inspectPdf(sourcePdfPath, workDir);
        const candidatePdf = await inspectPdf(candidatePdfPath, workDir);
        const libreOfficeStats = docxStats(new Uint8Array(await readFile(roundTripPath)));
        reportResult.libreOffice = {
          status: "pass",
          testedAt: new Date().toISOString(),
          pages: candidatePdf.pages,
          summary: structureSummary(libreOfficeStats),
          previews: await renderPdfPreviews(candidatePdfPath, fixture, "libreoffice", candidatePdf.pages),
        };
        assertStructure(sourceStats, libreOfficeStats, `${fixture} LibreOffice round trip`);
        assertPdf(wordReference ?? sourcePdf, candidatePdf, `${fixture} LibreOffice PDF`);

        let googleSummary = "";
        if (googleToken) {
          activeEngine = "googleDocs";
          const imported = await googleImport(googleToken, candidatePath, `WordInWeb interop smoke - ${fixture}`);
          try {
            const googleDocx = await googleExport(googleToken, imported.id,
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            const googlePdfBytes = await googleExport(googleToken, imported.id, "application/pdf");
            const googlePdfPath = join(workDir, `${basename(fixture, ".docx")}-google.pdf`);
            await writeFile(googlePdfPath, googlePdfBytes);
            const googleStats = docxStats(googleDocx);
            const googlePdf = await inspectPdf(googlePdfPath, workDir);
            reportResult.googleDocs = {
              status: "pass",
              testedAt: new Date().toISOString(),
              pages: googlePdf.pages,
              summary: structureSummary(googleStats),
              previews: await renderPdfPreviews(googlePdfPath, fixture, "google-docs", googlePdf.pages),
            };
            assertStructure(sourceStats, googleStats, `${fixture} Google Docs round trip`);
            assertPdf(wordReference ?? sourcePdf, googlePdf, `${fixture} Google Docs PDF`);
            googleSummary = `; Google Docs ${structureSummary(googleStats)}, ${googlePdf.pages} pages`;
          } finally {
            await googleDelete(googleToken, imported.id);
          }
        }

        console.log(`PASS ${fixture}: source ${structureSummary(sourceStats)}; LibreOffice ${structureSummary(libreOfficeStats)}, ${candidatePdf.pages} pages${googleSummary}`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        failures.push(`${fixture}: ${message}`);
        reportResult[activeEngine] = {
          ...reportResult[activeEngine],
          status: "fail",
          error: message,
        };
        console.error(`FAIL ${failures.at(-1)}`);
      }
      reportResults.push(reportResult);
    }
    await writeReportManifest(reportResults);
    run(process.execPath, [resolve("scripts/parity-render-report.mjs")]);
    await copyFile(resolve("parity/out/report.html"), resolve("apps/demo/public/report.html"));
  } finally {
    if (!keepArtifacts && failures.length === 0) await rm(workDir, { recursive: true, force: true });
  }

  if (failures.length > 0) {
    console.error(`\n${failures.length} compatibility failure(s):\n- ${failures.join("\n- ")}`);
    process.exitCode = 1;
  } else {
    console.log(`\nAll cross-editor layout smoke checks passed${includeGoogle ? " in LibreOffice and Google Docs" : " in LibreOffice"}.`);
  }
}

await main();
