import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { DocxDocument, insertModel3DAt } from "../../../../wordinweb/packages/core/dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const output = join(here, "../public/fixtures/model3d-cube.docx");
const poster = readFileSync(join(here, "assets/model3d-cube-poster.png"));

function cubeGlb() {
  const positions = new Float32Array([
    -1, -1, -1,  1, -1, -1,  1, 1, -1, -1, 1, -1,
    -1, -1,  1,  1, -1,  1,  1, 1,  1, -1, 1,  1,
  ]);
  const indices = new Uint16Array([
    0, 1, 2, 0, 2, 3, 4, 6, 5, 4, 7, 6,
    0, 4, 5, 0, 5, 1, 1, 5, 6, 1, 6, 2,
    2, 6, 7, 2, 7, 3, 3, 7, 4, 3, 4, 0,
  ]);
  const positionBytes = Buffer.from(positions.buffer);
  const indexBytes = Buffer.from(indices.buffer);
  const binary = Buffer.concat([positionBytes, indexBytes]);
  const gltf = {
    asset: { version: "2.0", generator: "WordInWeb demo fixture" },
    scene: 0,
    scenes: [{ nodes: [0] }],
    nodes: [{ mesh: 0, rotation: [0.163176, 0.336824, 0.059391, 0.925417] }],
    meshes: [{ primitives: [{ attributes: { POSITION: 0 }, indices: 1 }] }],
    buffers: [{ byteLength: binary.length }],
    bufferViews: [
      { buffer: 0, byteOffset: 0, byteLength: positionBytes.length, target: 34962 },
      { buffer: 0, byteOffset: positionBytes.length, byteLength: indexBytes.length, target: 34963 },
    ],
    accessors: [
      { bufferView: 0, componentType: 5126, count: 8, type: "VEC3", min: [-1, -1, -1], max: [1, 1, 1] },
      { bufferView: 1, componentType: 5123, count: indices.length, type: "SCALAR" },
    ],
  };
  const json = Buffer.from(JSON.stringify(gltf));
  const jsonPadding = (4 - (json.length % 4)) % 4;
  const binaryPadding = (4 - (binary.length % 4)) % 4;
  const jsonChunk = Buffer.concat([json, Buffer.alloc(jsonPadding, 0x20)]);
  const binaryChunk = Buffer.concat([binary, Buffer.alloc(binaryPadding)]);
  const glb = Buffer.alloc(12 + 8 + jsonChunk.length + 8 + binaryChunk.length);
  glb.writeUInt32LE(0x46546c67, 0);
  glb.writeUInt32LE(2, 4);
  glb.writeUInt32LE(glb.length, 8);
  glb.writeUInt32LE(jsonChunk.length, 12);
  glb.writeUInt32LE(0x4e4f534a, 16);
  jsonChunk.copy(glb, 20);
  const binaryHeader = 20 + jsonChunk.length;
  glb.writeUInt32LE(binaryChunk.length, binaryHeader);
  glb.writeUInt32LE(0x004e4942, binaryHeader + 4);
  binaryChunk.copy(glb, binaryHeader + 8);
  return glb;
}

const base = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{
    children: [
      new Paragraph({ text: "Native 3D model", heading: HeadingLevel.TITLE }),
      new Paragraph({
        children: [
          new TextRun({
            text: "This fixture carries a real GLB mesh, an Office 2019 model3d relationship, and a compatibility poster.",
            color: "44546A",
          }),
        ],
      }),
      new Paragraph({ children: [new TextRun("Interactive 3D asset:")] }),
    ],
  }],
});

const doc = DocxDocument.load(new Uint8Array(await Packer.toBuffer(base)));
let anchor;
for (const block of doc.sections[0].blocks) {
  if (block.type !== "paragraph") continue;
  for (const child of block.children) {
    const runs = child.type === "run" ? [child] : child.runs;
    for (const run of runs) {
      const text = run.content.find((content) => content.kind === "text" && content.text === "Interactive 3D asset:");
      if (text?.kind === "text" && text.srcT) anchor = text.srcT;
    }
  }
}
if (!anchor) throw new Error("3D fixture anchor text was not parsed");
if (!insertModel3DAt(doc, anchor, {
  data: new Uint8Array(cubeGlb()),
  poster: new Uint8Array(poster),
  alt: "Blue isometric cube 3D model",
})) throw new Error("3D model insertion failed");

const saved = doc.save();
writeFileSync(output, saved);
console.log("Wrote", output, saved.length, "bytes");
