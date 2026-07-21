import JSZip from "jszip";
import * as XLSX from "xlsx";

const clean = (value: unknown) => String(value ?? "").trim();
const headerKey = (value: unknown) =>
  clean(value)
    .normalize("NFKC")
    .replace(/[^0-9A-Za-z가-힣ㄱ-ㅎㅏ-ㅣ]/g, "");
const normalizedBarcode = (value: unknown) => clean(value).replace(/\.0$/, "").replace(/\D/g, "");

type BarcodeRow = {
  excelRow: number;
  itemNo: string;
  barcode: string;
  productImageThumbUrl: string;
  productImageOriginalUrl: string;
};

type BarcodeSheetMap = {
  rows: BarcodeRow[];
  imageColumn?: number;
};

export type BarcodeWorkbookData = {
  barcodeImages: Record<string, string>;
  productImagesByItemNo: Record<string, ProductImageReference>;
  productImagesByBarcode: Record<string, ProductImageReference>;
};

export type ProductImageReference = {
  thumbUrl: string;
  originalUrl: string;
};

const findHeaderIndex = (headers: string[], candidates: string[]) =>
  headers.findIndex((header) => candidates.some((candidate) => header === candidate || header.includes(candidate)));

export async function extractProductImageReferences(source: Blob | ArrayBuffer) {
  const buffer = source instanceof Blob ? await source.arrayBuffer() : source;
  const workbook = XLSX.read(buffer, { type: "array" });
  const productImagesByItemNo: Record<string, ProductImageReference> = {};
  const productImagesByBarcode: Record<string, ProductImageReference> = {};

  workbook.SheetNames.forEach((sheetName) => {
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, defval: "" });
    const headerIndex = matrix.findIndex((row) => {
      const values = row.map(headerKey);
      return values.some((value) => value.includes("순번") || value.includes("물품코드") || value.includes("품목코드"))
        && values.some((value) => value.includes("바코드"))
        && values.some((value) => value.includes("썸네일") || value.includes("이미지원본") || value.includes("이미지링크"));
    });
    if (headerIndex === -1) return;
    const headers = matrix[headerIndex].map(headerKey);
    const itemNoColumn = findHeaderIndex(headers, ["순번", "물품코드", "품목코드"]);
    const barcodeColumn = findHeaderIndex(headers, ["바코드"]);
    const thumbColumn = findHeaderIndex(headers, ["썸네일링크", "상품이미지미리보기링크", "미리보기링크", "thumbnail"]);
    const originalColumn = findHeaderIndex(headers, ["이미지원본링크", "원본이미지링크", "상품이미지원본링크", "상품이미지링크", "이미지링크"]);
    if (itemNoColumn === -1 || barcodeColumn === -1 || (thumbColumn === -1 && originalColumn === -1)) return;

    matrix.slice(headerIndex + 1).forEach((row) => {
      const itemNo = clean(row[itemNoColumn]);
      const barcode = normalizedBarcode(row[barcodeColumn]);
      const thumbUrl = thumbColumn === -1 ? "" : clean(row[thumbColumn]);
      const originalUrl = originalColumn === -1 ? "" : clean(row[originalColumn]);
      if (!itemNo || (!thumbUrl && !originalUrl)) return;
      const image = { thumbUrl: thumbUrl || originalUrl, originalUrl: originalUrl || thumbUrl };
      if (!productImagesByItemNo[itemNo]) productImagesByItemNo[itemNo] = image;
      if (barcode && !productImagesByBarcode[barcode]) productImagesByBarcode[barcode] = image;
    });
  });

  return { productImagesByItemNo, productImagesByBarcode };
}

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
    const values = row.map(headerKey);
    return values.some((value) => value.includes("순번") || value.includes("물품코드") || value.includes("품목코드"))
      && values.some((value) => value.includes("물품명") || value.includes("품목명") || value.includes("상품명"))
      && values.some((value) => value.includes("바코드"));
  });
  if (headerIndex === -1) return null;
  const headers = matrix[headerIndex].map(headerKey);
  const itemNoColumn = headers.findIndex((header) => header.includes("순번") || header.includes("물품코드") || header.includes("품목코드"));
  const barcodeColumn = headers.findIndex((header) => header.includes("바코드") && !header.includes("이미지"));
  const imageColumn = headers.findIndex((header) => header.includes("바코드") && header.includes("이미지"));
  const productThumbColumn = headers.findIndex((header) => ["썸네일링크", "상품이미지미리보기링크", "미리보기링크", "thumbnail"].includes(header));
  const productOriginalColumn = headers.findIndex((header) => ["이미지원본링크", "원본이미지링크", "상품이미지원본링크", "상품이미지링크", "이미지링크"].includes(header));
  if (itemNoColumn === -1 || barcodeColumn === -1 || (imageColumn === -1 && productThumbColumn === -1 && productOriginalColumn === -1)) return null;
  const rows = matrix
    .slice(headerIndex + 1)
    .map((row, index) => ({
      excelRow: headerIndex + index + 2,
      itemNo: clean(row[itemNoColumn]),
      barcode: normalizedBarcode(row[barcodeColumn]),
      productImageThumbUrl: productThumbColumn === -1 ? "" : clean(row[productThumbColumn]),
      productImageOriginalUrl: productOriginalColumn === -1 ? "" : clean(row[productOriginalColumn]),
    }))
    .filter((row) => row.itemNo);
  return { rows, imageColumn: imageColumn === -1 ? undefined : imageColumn + 1 };
}

export async function extractBarcodeWorkbookData(source: Blob | ArrayBuffer): Promise<BarcodeWorkbookData> {
  const buffer = source instanceof Blob ? await source.arrayBuffer() : source;
  const workbook = XLSX.read(buffer, { type: "array" });
  const zip = await JSZip.loadAsync(buffer);
  const barcodeImages: Record<string, string> = {};
  const productImagesByItemNo: Record<string, ProductImageReference> = {};
  const productImagesByBarcode: Record<string, ProductImageReference> = {};
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const sheetPathByName = workbookXml && workbookRelsXml ? workbookSheetPaths(workbookXml, workbookRelsXml) : new Map<string, string>();

  for (const [sheetIndex, sheetName] of workbook.SheetNames.entries()) {
    const barcodeSheet = barcodeRowsFromSheet(workbook.Sheets[sheetName]);
    if (!barcodeSheet?.rows.length) continue;
    const barcodeRowByRow = new Map(barcodeSheet.rows.map((row) => [row.excelRow, row]));
    barcodeSheet.rows.forEach((row) => {
      const image = {
        thumbUrl: row.productImageThumbUrl || row.productImageOriginalUrl,
        originalUrl: row.productImageOriginalUrl || row.productImageThumbUrl,
      };
      if (!image.thumbUrl && !image.originalUrl) return;
      if (!productImagesByItemNo[row.itemNo]) productImagesByItemNo[row.itemNo] = image;
      if (row.barcode && !productImagesByBarcode[row.barcode]) productImagesByBarcode[row.barcode] = image;
    });

    if (!barcodeSheet.imageColumn) continue;

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
      const row = barcodeRowByRow.get(anchorRow);
      if (!row?.barcode) continue;
      const blip = anchor.getElementsByTagName("a:blip")[0] ?? anchor.getElementsByTagName("blip")[0];
      const imageRid = blip?.getAttribute("r:embed") ?? blip?.getAttribute("embed");
      const imagePath = imageRid ? drawingRels.get(imageRid) : "";
      if (!imagePath) continue;
      const imageBuffer = await zip.file(imagePath)?.async("arraybuffer");
      if (!imageBuffer) continue;
      const dataUrl = arrayBufferToDataUrl(imageBuffer, mimeFromPath(imagePath));
      if (!barcodeImages[row.barcode]) barcodeImages[row.barcode] = dataUrl;
    }
  }

  return { barcodeImages, productImagesByItemNo, productImagesByBarcode };
}

export async function extractBarcodeImages(source: Blob | ArrayBuffer) {
  return (await extractBarcodeWorkbookData(source)).barcodeImages;
}
