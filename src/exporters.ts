import JSZip from "jszip";
import * as XLSX from "xlsx";
import { downloadBlob, safeFilePart } from "./logic";
import type { AppSettings, BackupPayload, Region, SurveyItem, SurveyPhoto, SurveyStore } from "./types";

const stamp = () => new Date().toISOString().slice(0, 10).replaceAll("-", "");
const stampTime = () => new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");

function photoOf(photos: SurveyPhoto[], type: SurveyPhoto["type"], item?: SurveyItem, store?: SurveyStore) {
  return photos.find((photo) => photo.type === type && (item ? photo.itemId === item.id : true) && (store ? photo.storeId === store.id : true));
}

function districtFromAddress(address: string) {
  const cleaned = address.replace(/^\(\d{5}\)\s*/, "").trim();
  const match = cleaned.match(/(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)[^\s]*\s+([^\s]+(?:시|군|구))/);
  return match?.[1] ?? "";
}

function priceJudgmentOf(item: SurveyItem) {
  if (item.basePrice === null || item.normalPrice === null) return "";
  if (item.normalPrice === item.basePrice) return "동일";
  return item.normalPrice > item.basePrice ? "고가" : "저가";
}

export async function exportRegionExcel(region: string, items: SurveyItem[]) {
  const group = [
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "기준 정보",
    "실물 확인가능",
    "실물 확인가능",
    "실물 확인가능",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "가격 판단",
    "정상진열 안될 시",
    "정상진열 안될 시",
    "정상진열 안될 시",
    "정상진열 안될 시",
    "특이사항",
    "조사일",
    "마트정보",
    "마트정보",
    "마트정보",
    "업체정보",
    "업체정보",
  ];
  const headers = [
    "순번",
    "업체명",
    "업체연락처",
    "마트명",
    "바코드",
    "물품명",
    "규격",
    "기준가격",
    "정상진열",
    "규격일치",
    "바코드일치",
    "정상가격",
    "할인가격",
    "시작",
    "종료",
    "기간구분",
    "가격 판단",
    "바코드등록여부",
    "미판매",
    "미진열",
    "비정상",
    "특이사항",
    "조사일",
    "광역",
    "시군구",
    "주소",
    "담당자",
    "업체 연락처",
  ];
  const rows = items.map((item) => [
    item.itemNo,
    item.companyName,
    item.companyTel,
    item.martName,
    item.barcode,
    item.productName,
    item.spec,
    item.basePrice ?? "",
    item.normalDisplay,
    item.specMatch,
    item.barcodeMatch,
    item.normalPrice ?? "",
    item.discountPrice ?? "",
    item.discountStartDate,
    item.discountEndDate,
    `${item.discountType.replace("구두", "") === "①" ? "1" : item.discountType.replace("구두", "") === "②" ? "2" : item.discountType.replace("구두", "")}${item.discountOral || item.discountType.includes("구두") ? "구두확인" : ""}`,
    priceJudgmentOf(item),
    item.barcodeRegistered,
    item.abnormalStatus === "미판매" ? "O" : item.abnormalStatus ? "-" : "",
    item.abnormalStatus === "미진열" ? "O" : item.abnormalStatus ? "-" : "",
    item.abnormalDisplay === "O" ? "O" : "",
    item.memo,
    item.status === "미조사" ? "" : item.surveyDate,
    item.city ?? "",
    districtFromAddress(item.storeAddress),
    item.storeAddress,
    item.companyManager ?? "",
    item.companyTel || item.martTel || "",
  ]);
  const ws = XLSX.utils.aoa_to_sheet([group, headers, ...rows]);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
    { s: { r: 0, c: 8 }, e: { r: 0, c: 10 } },
    { s: { r: 0, c: 11 }, e: { r: 0, c: 16 } },
    { s: { r: 0, c: 17 }, e: { r: 0, c: 20 } },
    { s: { r: 0, c: 23 }, e: { r: 0, c: 25 } },
    { s: { r: 0, c: 26 }, e: { r: 0, c: 27 } },
  ];
  ws["!cols"] = headers.map((header) => ({ wch: Math.max(10, header.length + 4) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "조사결과");
  const buffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  await downloadBlob(new Blob([buffer], { type: "application/octet-stream" }), `price_survey_${safeFilePart(region)}_${stamp()}.xlsx`);
}

export async function exportRegionZip(region: string, stores: SurveyStore[], items: SurveyItem[], photos: SurveyPhoto[]) {
  const zip = new JSZip();
  items.forEach((item) => {
    const store = stores.find((candidate) => candidate.id === item.storeId);
    const front = store ? photoOf(photos, "STORE_FRONT", undefined, store) : undefined;
    const display = photoOf(photos, "PRODUCT_DISPLAY", item);
    const info = photoOf(photos, "PRODUCT_INFO_BARCODE", item);
    const pos = photoOf(photos, "POS_RECEIPT", item);
    const name = safeFilePart(item.itemNo || item.productName);
    if (front) zip.file(`${name}.1.jpg`, front.blob);
    if (display) zip.file(`${name}.2.jpg`, display.blob);
    if (info) zip.file(`${name}.3.jpg`, info.blob);
    if (pos) zip.file(`${name}.4.jpg`, pos.blob);
  });
  await downloadBlob(await zip.generateAsync({ type: "blob", mimeType: "application/zip" }), `price_photos_${safeFilePart(region)}_${stamp()}.zip`);
}

const blobToDataUrl = (blob: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });

export async function exportBackup(
  region: string | undefined,
  regions: Region[],
  stores: SurveyStore[],
  items: SurveyItem[],
  photos: SurveyPhoto[],
  settings: AppSettings,
  barcodeIndex: Record<string, string> = {},
) {
  const photoPayload = await Promise.all(
    photos.map(async ({ blob, ...photo }) => ({
      ...photo,
      dataUrl: await blobToDataUrl(blob),
    })),
  );
  const payload: BackupPayload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    scope: region ? "region" : "all",
    region,
    regions,
    stores,
    items,
    photos: photoPayload,
    barcodeIndex,
    settings,
  };
  const suffix = region ? safeFilePart(region) : "전체";
  await downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }), `price_backup_${suffix}_${stampTime()}.json`);
}

export async function dataUrlToBlob(dataUrl: string) {
  const response = await fetch(dataUrl);
  return await response.blob();
}
