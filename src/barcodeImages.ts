import JSZip from "jszip";
import * as XLSX from "xlsx";

const clean = (value: unknown) => String(value ?? "").trim();

type BarcodeRow = {
  excelRow: number;
  itemNo: string;
};

type BarcodeSheetMap = {
  rows: BarcodeRow[];
  imageColumn: number;
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

function workbookSheetPaths(workbookXml: string, workbookRelsXml: string) {
  const doc = parseXml(workbookXml);
  const workbookRels = relationMap(workbookRelsXml, "xl/workbook.xml");
  const map = new Map<string, string>();
  Array.from(doc.getElementsByTagName("sheet")).forEach((node) => {
    const name = node.getAttribute("name");
    const rid = node.getAttribute("r:id") ?? node.getAttribute("id");
    const sheetPath = rid ? workbookRels.get(rid) : "";
    if (name && sheetPath) map.set(name, sheetPath);
  });
  return map;
}

function barcodeRowsFromSheet(sheet: XLSX.WorkSheet): BarcodeSheetMap | null {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
  const headerIndex = matrix.findIndex((row) => {
    const values = row.map(clean);
    return values.includes("순번") && values.includes("물품명") && values.includes("바코드");
  });
  if (headerIndex === -1) return null;
  const headers = matrix[headerIndex].map(clean);
  const itemNoColumn = headers.findIndex((header) => header === "순번");
  const barcodeColumn = headers.findIndex((header) => header === "바코드");
  const imageColumn = headers.findIndex((header) => header === "바코드 이미지");
  if (itemNoColumn === -1 || barcodeColumn === -1 || imageColumn === -1) return null;
  const rows = matrix
    .slice(headerIndex + 1)
    .map((row, index) => ({ excelRow: headerIndex + index + 2, itemNo: clean(row[itemNoColumn]) }))
    .filter((row) => row.itemNo);
  return { rows, imageColumn: imageColumn + 1 };
}

export async function extractBarcodeImages(source: Blob | ArrayBuffer) {
  const buffer = source instanceof Blob ? await source.arrayBuffer() : source;
  const workbook = XLSX.read(buffer, { type: "array" });
  const zip = await JSZip.loadAsync(buffer);
  const index: Record<string, string> = {};
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const sheetPathByName = workbookXml && workbookRelsXml ? workbookSheetPaths(workbookXml, workbookRelsXml) : new Map<string, string>();

  for (const [sheetIndex, sheetName] of workbook.SheetNames.entries()) {
    const barcodeSheet = barcodeRowsFromSheet(workbook.Sheets[sheetName]);
    if (!barcodeSheet?.rows.length) continue;
    const itemNoByRow = new Map(barcodeSheet.rows.map((row) => [row.excelRow, row.itemNo]));

    const sheetPath = sheetPathByName.get(sheetName) ?? `xl/worksheets/sheet${sheetIndex + 1}.xml`;
    const sheetXml = await zip.file(sheetPath)?.async("string");
    const sheetFileName = sheetPath.slice(sheetPath.lastIndexOf("/") + 1);
    const sheetRelsXml = await zip.file(`${sheetPath.slice(0, sheetPath.lastIndexOf("/"))}/_rels/${sheetFileName}.rels`)?.async("string");
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
      const colText = anchor.getElementsByTagName("xdr:col")[0]?.textContent ?? anchor.getElementsByTagName("col")[0]?.textContent ?? "";
      const anchorRow = Number(rowText) + 1;
      const anchorColumn = Number(colText) + 1;
      if (!Number.isFinite(anchorRow) || anchorColumn !== barcodeSheet.imageColumn) continue;
      const itemNo = itemNoByRow.get(anchorRow) ?? "";
      if (!itemNo || index[itemNo]) continue;
      const blip = anchor.getElementsByTagName("a:blip")[0] ?? anchor.getElementsByTagName("blip")[0];
      const imageRid = blip?.getAttribute("r:embed") ?? blip?.getAttribute("embed");
      const imagePath = imageRid ? drawingRels.get(imageRid) : "";
      if (!imagePath) continue;
      const imageBuffer = await zip.file(imagePath)?.async("arraybuffer");
      if (!imageBuffer) continue;
      index[itemNo] = arrayBufferToDataUrl(imageBuffer, mimeFromPath(imagePath));
    }
  }

  return index;
}
