/**
 * Shared desktop-Word export and raster helpers.
 *
 * Both DOCX parity gates — scripts/word-download-parity.mjs (saved fixture vs
 * cached Word reference) and scripts/edit-roundtrip-parity.mjs (edited document
 * in Word vs the same document on the web) — drive the same pipeline: hand a
 * DOCX to desktop Microsoft Word over AppleScript, take the PDF it exports,
 * rasterize with pdftoppm at 192 DPI, and count mismatched pixels.
 *
 * Inputs and outputs are staged inside Word's own container so the run does not
 * prompt for file access on every document. Microsoft Word needs Full Disk
 * Access on macOS.
 */

import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { basename, join } from "node:path";
import { unzipSync } from "fflate";

/** Word's own Documents container. Staging here avoids a sandbox prompt per file. */
export const wordIoDir = join(homedir(), "Library/Containers/com.microsoft.Word/Data/Documents/WordInWebParity");

/** Device pixels per CSS px on the web side; pdftoppm dpi is 96 * SCALE. */
export const SCALE = 2;
export const DPI = 96 * SCALE;

export function sha256(path) {
  return execFileSync("shasum", ["-a", "256", path], { encoding: "utf8" }).trim().split(/\s+/)[0];
}

/** Hash of a DOCX's *contents*, ignoring zip framing (order, timestamps). */
export function packageSha256(path) {
  const hash = createHash("sha256");
  const files = unzipSync(readFileSync(path));
  for (const name of Object.keys(files).sort()) {
    const bytes = files[name];
    hash.update(name);
    hash.update("\0");
    hash.update(String(bytes.length));
    hash.update("\0");
    hash.update(bytes);
  }
  return hash.digest("hex");
}

export function pdfInfo(path) {
  const text = execFileSync("pdfinfo", [path], { encoding: "utf8" });
  const creator = text.match(/^Creator:\s*(.*)$/m)?.[1]?.trim() ?? "";
  const pages = Number(text.match(/^Pages:\s*(\d+)$/m)?.[1] ?? 0);
  // Word occasionally preserves an empty Creator field from the source
  // document. The AppleScript export in this process establishes candidate
  // provenance; reject any non-empty metadata that names another producer.
  if (creator && creator !== "Microsoft Word") {
    throw new Error(`${path} is not a Microsoft Word PDF (Creator=${creator})`);
  }
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`Could not read page count from ${path}`);
  return { creator, pages };
}

/** Page PNGs written by pdftoppm under `prefix`, in page order. */
export function pngs(dir, prefix) {
  return readdirSync(dir)
    .filter((name) => name.startsWith(`${prefix}-`) && name.endsWith(".png"))
    .sort((a, b) => Number(a.match(/-(\d+)\.png$/)?.[1]) - Number(b.match(/-(\d+)\.png$/)?.[1]))
    .map((name) => join(dir, name));
}

/** Rasterize `pdf` into `cacheDir` once; later runs reuse the completed set. */
export function ensureRasters(pdf, cacheDir, prefix, expectedPages) {
  const complete = join(cacheDir, ".complete");
  if (!existsSync(complete) || pngs(cacheDir, prefix).length !== expectedPages) {
    const temp = `${cacheDir}.tmp-${process.pid}`;
    rmSync(temp, { recursive: true, force: true });
    mkdirSync(temp, { recursive: true });
    execFileSync("pdftoppm", ["-r", String(DPI), "-png", pdf, join(temp, prefix)], { stdio: "inherit" });
    writeFileSync(join(temp, ".complete"), "");
    rmSync(cacheDir, { recursive: true, force: true });
    renameSync(temp, cacheDir);
  }
  return cacheDir;
}

const escapeAppleScript = (value) => value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');

/**
 * Open `docx` in desktop Word and save it as PDF at `destination`.
 *
 * PDFs are cached by the source package hash under `cacheDir`, so re-running a
 * scenario whose edit produced identical content costs no Word round trip.
 * Word's repair prompt is a modal dialog that never answers AppleScript, so a
 * damaged package surfaces here as a throw (open error, or the 600s timeout) —
 * callers treat that as a failed open, not as a missing PDF.
 */
export function exportWithWord({ name, docx, destination, packageHash, cacheDir }) {
  const cachedPdf = join(cacheDir, `${packageHash}.pdf`);
  if (existsSync(cachedPdf) && statSync(cachedPdf).size > 0) {
    pdfInfo(cachedPdf);
    copyFileSync(cachedPdf, destination);
    console.log(`Reused ${cachedPdf}`);
    return { reused: true };
  }
  const stagedDocx = join(wordIoDir, `${name}-website.docx`);
  const stagedPdf = join(wordIoDir, `${name}-website-word.pdf`);
  rmSync(stagedDocx, { force: true });
  rmSync(stagedPdf, { force: true });
  copyFileSync(docx, stagedDocx);
  const script = `with timeout of 600 seconds\n` +
    `tell application "Microsoft Word"\n` +
    `  try\n` +
    `    close document "${escapeAppleScript(basename(stagedDocx))}" saving no\n` +
    `  end try\n` +
    `  open file name "${escapeAppleScript(stagedDocx)}"\n` +
    `  repeat with attempt from 1 to 120\n` +
    `    if exists document "${escapeAppleScript(basename(stagedDocx))}" then exit repeat\n` +
    `    delay 1\n` +
    `  end repeat\n` +
    `  if not (exists document "${escapeAppleScript(basename(stagedDocx))}") then error "Word did not finish opening ${escapeAppleScript(basename(stagedDocx))}"\n` +
    `  set candidateDocument to document "${escapeAppleScript(basename(stagedDocx))}"\n` +
    `  delay 5\n` +
    `  save as candidateDocument file name "${escapeAppleScript(stagedPdf)}" file format format PDF\n` +
    `  close candidateDocument saving no\n` +
    `end tell\n` +
    `end timeout`;
  execFileSync("osascript", ["-e", script], { timeout: 610_000, stdio: "inherit" });
  if (!existsSync(stagedPdf) || statSync(stagedPdf).size === 0) throw new Error(`Word export failed for ${name}`);
  pdfInfo(stagedPdf);
  copyFileSync(stagedPdf, destination);
  copyFileSync(stagedPdf, cachedPdf);
  console.log(`Wrote ${stagedPdf}`);
  return { reused: false };
}

/**
 * Mismatched-pixel count between two PNGs, measured in a browser page so both
 * sides decode through the same image pipeline. Sizes are unioned onto a white
 * canvas, so a page that differs in size counts the surplus as mismatch.
 */
export async function comparePngs(page, reference, candidate) {
  const [referenceBytes, candidateBytes] = [readFileSync(reference), readFileSync(candidate)];
  return page.evaluate(async ({ referenceData, candidateData }) => {
    const load = (data) => new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error("PNG decode failed"));
      image.src = `data:image/png;base64,${data}`;
    });
    const [a, b] = await Promise.all([load(referenceData), load(candidateData)]);
    const width = Math.max(a.width, b.width);
    const height = Math.max(a.height, b.height);
    const pixels = (image) => {
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d", { willReadFrequently: true });
      context.fillStyle = "white";
      context.fillRect(0, 0, width, height);
      context.drawImage(image, 0, 0);
      return context.getImageData(0, 0, width, height).data;
    };
    const aa = pixels(a);
    const bb = pixels(b);
    let mismatchedPixels = 0;
    for (let offset = 0; offset < aa.length; offset += 4) {
      const delta = Math.abs(aa[offset] - bb[offset])
        + Math.abs(aa[offset + 1] - bb[offset + 1])
        + Math.abs(aa[offset + 2] - bb[offset + 2]);
      if (delta > 90) mismatchedPixels++;
    }
    return { width, height, pixels: width * height, mismatchedPixels };
  }, {
    referenceData: referenceBytes.toString("base64"),
    candidateData: candidateBytes.toString("base64"),
  });
}
