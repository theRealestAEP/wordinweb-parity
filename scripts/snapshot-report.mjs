import { execFile } from "node:child_process";
import { access, copyFile, mkdir, readFile } from "node:fs/promises";
import { basename, join } from "node:path";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const sourceDirectory = "parity/out";
const publicDirectory = "apps/demo/public";
const sourceReport = join(sourceDirectory, "report.html");
const publicReport = join(publicDirectory, "report.html");

const html = await readFile(sourceReport, "utf8");
const images = [
  ...new Set(
    [...html.matchAll(/(?:href|data-png)="([^"]+\.png)"/g)].map(
      ([, file]) => file,
    ),
  ),
];

await mkdir(publicDirectory, { recursive: true });
await copyFile(sourceReport, publicReport);

let nextImage = 0;

async function convertImages() {
  while (nextImage < images.length) {
    const image = images[nextImage++];
    if (basename(image) !== image) {
      throw new Error(`Unexpected report image path: ${image}`);
    }

    const sourceImage = join(sourceDirectory, image);
    const publicImage = join(publicDirectory, image);
    try {
      await access(sourceImage);
    } catch {
      await access(publicImage);
      continue;
    }

    await execFileAsync("magick", [
      sourceImage,
      "-resize",
      "1800x>",
      "-strip",
      "-colors",
      "128",
      `PNG8:${publicImage}`,
    ]);
  }
}

await Promise.all(Array.from({ length: 4 }, convertImages));
console.log(`Snapshotted report and ${images.length} review images.`);
