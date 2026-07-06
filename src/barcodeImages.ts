import JSZip from "jszip";
import * as XLSX from "xlsx";

const clean = (value: unknown) => String(value ?? "").trim();
const onlyDigits = (value: string) => value.replace(/\D/g, "");

type BarcodeRow = {
  excelRow: number;
  barcode: string;
};

const mimeFromPath = (path: string) => {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".webp")) return "image/webp";
  return "image/jpeg";
};

const normalizeZipPath = (base: string, target: string) => {
  const baseParts = base.split("/");
  baseParts.pop();
  const parts = `${baseParts.join("/")}/${target}`.split("/");
  const out: string[] = [];
  parts.forEach((part) => {
    if (!part || part === ".") return;
    if (part === "..") out.pop();
    else out.push(part);
  });
  return out.join("/");
};

const arrayBufferToDataUrl = (buffer: ArrayBuffer, mimeType: string) => {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  }
  return `data:${mimeType};base64,${btoa(binary)}`;
};

function parseXml(text: string) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function relationMap(xml: string, basePath: string) {
  const doc = parseXml(xml);
  const map = new Map<string, string>();
  Array.from(doc.getElementsByTagName("Relationship")).forEach((node) => {
    const id = node.getAttribute("Id");
    const target = node.getAttribute("Target");
    if (!id || !target) return;
    map.set(id, normalizeZipPath(basePath, target));
  });
  return map;
}

function barcodeRowsFromSheet(sheet: XLSX.WorkSheet): BarcodeRow[] {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
  const headerIndex = matrix.findIndex((row) => row.map(clean).some((value) => value.includes("바코드")));
  if (headerIndex === -1) return [];
  const headers = matrix[headerIndex].map(clean);
  const barcodeColumn = headers.findIndex((header) => header.includes("바코드"));
  if (barcodeColumn === -1) return [];
  return matrix
    .slice(headerIndex + 1)
    .map((row, index) => ({ excelRow: headerIndex + index + 1, barcode: onlyDigits(clean(row[barcodeColumn])) }))
    .filter((row) => row.barcode);
}

function nearestBarcode(rows: BarcodeRow[], anchorRow: number) {
  let best: BarcodeRow | undefined;
  let bestDistance = Number.POSITIVE_INFINITY;
  rows.forEach((row) => {
    const distance = Math.abs(row.excelRow - anchorRow);
    if (distance < bestDistance && distance <= 2) {
      best = row;
      bestDistance = distance;
    }
  });
  return best?.barcode ?? "";
}

export async function extractBarcodeImages(source: Blob | ArrayBuffer) {
  const buffer = source instanceof Blob ? await source.arrayBuffer() : source;
  const workbook = XLSX.read(buffer, { type: "array" });
  const zip = await JSZip.loadAsync(buffer);
  const index: Record<string, string> = {};

  for (const [sheetIndex, sheetName] of workbook.SheetNames.entries()) {
    const rows = barcodeRowsFromSheet(workbook.Sheets[sheetName]);
    if (!rows.length) continue;

    const sheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
    const sheetXml = await zip.file(sheetPath)?.async("string");
    const sheetRelsXml = await zip.file(`xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`)?.async("string");
    if (!sheetXml || !sheetRelsXml) continue;

    const sheetDoc = parseXml(sheetXml);
    const drawingNode = Array.from(sheetDoc.getElementsByTagName("drawing"))[0];
    const drawingRid = drawingNode?.getAttribute("r:id");
    if (!drawingRid) continue;

    const sheetRels = relationMap(sheetRelsXml, sheetPath);
    const drawingPath = sheetRels.get(drawingRid);
    if (!drawingPath) continue;

    const drawingXml = await zip.file(drawingPath)?.async("string");
    const drawingRelsXml = await zip.file(`${drawingPath.slice(0, drawingPath.lastIndexOf("/"))}/_rels/${drawingPath.slice(drawingPath.lastIndexOf("/") + 1)}.rels`)?.async("string");
    if (!drawingXml || !drawingRelsXml) continue;

    const drawingRels = relationMap(drawingRelsXml, drawingPath);
    const drawingDoc = parseXml(drawingXml);
    const anchors = [
      ...Array.from(drawingDoc.getElementsByTagName("xdr:twoCellAnchor")),
      ...Array.from(drawingDoc.getElementsByTagName("xdr:oneCellAnchor")),
      ...Array.from(drawingDoc.getElementsByTagName("twoCellAnchor")),
      ...Array.from(drawingDoc.getElementsByTagName("oneCellAnchor")),
    ];

    for (const anchor of anchors) {
      const rowText = anchor.getElementsByTagName("xdr:row")[0]?.textContent ?? anchor.getElementsByTagName("row")[0]?.textContent ?? "";
      const anchorRow = Number(rowText) + 1;
      if (!Number.isFinite(anchorRow)) continue;
      const barcode = nearestBarcode(rows, anchorRow);
      if (!barcode || index[barcode]) continue;
      const blip = anchor.getElementsByTagName("a:blip")[0] ?? anchor.getElementsByTagName("blip")[0];
      const imageRid = blip?.getAttribute("r:embed") ?? blip?.getAttribute("embed");
      const imagePath = imageRid ? drawingRels.get(imageRid) : "";
      if (!imagePath) continue;
      const imageBuffer = await zip.file(imagePath)?.async("arraybuffer");
      if (!imageBuffer) continue;
      index[barcode] = arrayBufferToDataUrl(imageBuffer, mimeFromPath(imagePath));
    }
  }

  return index;
}
