import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import XLSX from "xlsx";
import JSZip from "jszip";

const source = process.argv[2];
const outputDir = process.argv[3] ?? "public/data/barcodes";

if (!source) {
  console.error("Usage: node scripts/extract-barcode-images.mjs <xlsx> [outputDir]");
  process.exit(1);
}

const workbook = XLSX.readFile(source);
const zip = await JSZip.loadAsync(fs.readFileSync(source));
fs.mkdirSync(outputDir, { recursive: true });

const readZipText = async (name) => {
  const file = zip.file(name);
  return file ? await file.async("string") : "";
};

const parseRels = (xml) => {
  const map = new Map();
  const pattern = /<Relationship\b([^>]+?)\/>/g;
  for (const match of xml.matchAll(pattern)) {
    const attrs = match[1];
    const id = attrs.match(/\bId="([^"]+)"/)?.[1];
    const target = attrs.match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) map.set(id, target);
  }
  return map;
};

const cleanBarcode = (value) => String(value ?? "").replace(/\D/g, "");
const extFromTarget = (target) => path.extname(target).toLowerCase() || ".png";

let copied = 0;
const index = {};

for (let sheetIndex = 0; sheetIndex < workbook.SheetNames.length; sheetIndex += 1) {
  const sheetName = workbook.SheetNames[sheetIndex];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
  const sheetNumber = sheetIndex + 1;
  const sheetRels = parseRels(await readZipText(`xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`));
  const drawingTarget = Array.from(sheetRels.values()).find((target) => target.includes("drawings/drawing"));
  if (!drawingTarget) continue;
  const drawingName = drawingTarget.split("/").at(-1);
  const drawingPath = `xl/drawings/${drawingName}`;
  const drawingXml = await readZipText(drawingPath);
  const drawingRels = parseRels(await readZipText(`xl/drawings/_rels/${drawingName}.rels`));

  const anchors = drawingXml.match(/<xdr:twoCellAnchor[\s\S]*?<\/xdr:twoCellAnchor>/g) ?? [];
  for (const anchor of anchors) {
    const col = Number(anchor.match(/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/)?.[1]);
    const row = Number(anchor.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/)?.[1]);
    if (col !== 5 || !Number.isFinite(row)) continue;
    const relId = anchor.match(/r:embed="([^"]+)"/)?.[1];
    const target = relId ? drawingRels.get(relId) : "";
    if (!target) continue;
    const barcode = cleanBarcode(rows[row]?.[4]);
    if (!barcode || index[barcode]) continue;
    const mediaPath = `xl/media/${target.split("/").at(-1)}`;
    const media = zip.file(mediaPath);
    if (!media) continue;
    const ext = extFromTarget(target);
    const fileName = `${barcode}${ext}`;
    fs.writeFileSync(path.join(outputDir, fileName), await media.async("nodebuffer"));
    index[barcode] = `data/barcodes/${fileName}`;
    copied += 1;
  }
}

fs.writeFileSync("public/data/barcode-index.json", `${JSON.stringify(index, null, 2)}\n`, "utf8");
console.log(`Extracted ${copied} barcode images to ${outputDir}`);
