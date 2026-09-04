import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const [artifactToolRoot, inputPptx, outputDir] = process.argv.slice(2);
if (!artifactToolRoot || !inputPptx || !outputDir) {
  throw new Error("Usage: render_pptx_images.mjs <node_modules> <input.pptx> <output-dir>");
}

const modulePath = path.join(
  artifactToolRoot,
  "@oai",
  "artifact-tool",
  "dist",
  "artifact_tool.mjs",
);
const { FileBlob, PresentationFile } = await import(pathToFileURL(modulePath).href);
const presentation = await PresentationFile.importPptx(await FileBlob.load(inputPptx));
await fs.mkdir(outputDir, { recursive: true });
for (const name of await fs.readdir(outputDir)) {
  if (/^slide-\d+\.png$/i.test(name)) await fs.unlink(path.join(outputDir, name));
}
for (const [index, slide] of presentation.slides.items.entries()) {
  const blob = await presentation.export({ slide, format: "png", scale: 1.5 });
  const filename = `slide-${String(index + 1).padStart(2, "0")}.png`;
  await fs.writeFile(path.join(outputDir, filename), new Uint8Array(await blob.arrayBuffer()));
}
