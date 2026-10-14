import * as XLSX from "xlsx";
import { now, today, uid } from "./db";
import type { Region, SurveyItem, SurveyStore } from "./types";

type Row = Record<string, unknown>;

const aliases: Record<string, string[]> = {
  itemNo: ["순번", "번호", "품목번호", "연번", "no"],
  companyName: ["업체명", "회사명", "제조회사", "제조사", "공급업체"],
  companyTel: ["업체연락처", "담당자연락처", "담당자 연락처", "연락처", "전화번호", "업체전화", "TEL"],
  companyManager: ["담당자", "담당", "업체담당자"],
  martName: ["마트명", "판매처", "조사처", "매장명", "방문지", "마트"],
  martTel: ["마트번호", "마트 연락처", "마트연락처", "판매처번호", "판매처연락처", "매장번호", "매장연락처"],
  barcode: ["바코드", "barcode", "상품코드"],
  productImageUrl: ["이미지원본링크", "원본이미지링크", "상품이미지원본링크", "썸네일링크", "상품이미지미리보기링크", "미리보기링크", "상품이미지링크", "이미지링크", "thumbnail"],
  productName: ["물품명", "품목명", "상품명", "제품명"],
  spec: ["규격", "용량", "단위"],
  basePrice: ["기준가격", "기준가", "가격", "기준단가"],
  region: ["권역", "담당권역"],
  city: ["광역", "시도"],
  district: ["시군구", "구군", "지역"],
  department: ["담당부서", "담당", "부서", "지회", "담당지회"],
  address: ["주소", "마트주소", "판매처주소", "매장주소"],
  finalAddress: ["최종주소", "최종 주소", "지도주소", "검색주소"],
  detailAddress: ["주소", "상세주소", "세부주소"],
  latitude: ["위도", "latitude", "lat", "y좌표", "y"],
  longitude: ["경도", "longitude", "lng", "lon", "x좌표", "x"],
};

const clean = (value: unknown) => String(value ?? "").trim();
const compact = (value: string) => value.toLowerCase().replace(/[\s_\-()[\]{}]/g, "");

// 엑셀 헤더는 정확히 일치하는 열을 우선 사용한다.
function pick(row: Row, key: keyof typeof aliases) {
  const headers = Object.keys(row).filter((header) => !header.startsWith("__"));
  for (const alias of aliases[key]) {
    const found = headers.find((header) => compact(header) === compact(alias));
    if (found) return clean(row[found]);
  }
  for (const alias of aliases[key]) {
    const found = headers.find((header) => compact(header).includes(compact(alias)));
    if (found) return clean(row[found]);
  }
  return "";
}

function numberOrNull(value: string) {
  if (!value.trim()) return null;
  const parsed = Number(value.replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function storeKey(region: string, name: string, address: string) {
  return `${region}__${name || "미상"}__${address || "주소없음"}`;
}

async function readWorkbook(source: Blob | ArrayBuffer) {
  const buffer = source instanceof Blob ? await source.arrayBuffer() : source;
  return XLSX.read(buffer, { type: "array" });
}

function dataSheetNames(workbook: XLSX.WorkBook) {
  if (workbook.SheetNames.includes("전체")) return ["전체"];
  return workbook.SheetNames.filter((name) => !["샘플", "sample"].includes(name.toLowerCase()));
}

export async function parseSurveyWorkbook(source: Blob | ArrayBuffer): Promise<{ regions: Region[]; stores: SurveyStore[]; items: SurveyItem[] }> {
  const workbook = await readWorkbook(source);
  const sheetNames = dataSheetNames(workbook);
  const storeMap = new Map<string, SurveyStore>();
  const regionMap = new Map<string, Region>();
  const districtMap = new Map<string, Set<string>>();
  const items: SurveyItem[] = [];
  const current = now();

  sheetNames.flatMap((name) => rowsFromSheet(workbook.Sheets[name], name)).forEach((row, index) => {
    const productName = pick(row, "productName");
    const barcode = pick(row, "barcode");
    if (!productName && !barcode) return;
    // 엑셀 `광역` 원문을 그대로 지역명으로 사용한다. 권역/담당부서는 생성하지 않는다.
    const rawCity = pick(row, "city");
    const cityName = rawCity || "미지정";
    const districtName = pick(row, "district");
    const resolved = { name: cityName, city: cityName, department: undefined };
    if (!districtMap.has(cityName)) districtMap.set(cityName, new Set<string>());
    if (districtName) districtMap.get(cityName)!.add(districtName);
    const martName = pick(row, "martName") || pick(row, "companyName") || "미상 방문지";
    const detailAddress = pick(row, "detailAddress") || pick(row, "address");
    const address = pick(row, "finalAddress") || detailAddress;
    const latitude = numberOrNull(pick(row, "latitude"));
    const longitude = numberOrNull(pick(row, "longitude"));
    const key = storeKey(resolved.name, martName, address);
    if (!storeMap.has(key)) {
      storeMap.set(key, {
        id: uid("store"),
        region: resolved.name,
        department: resolved.department,
        city: resolved.city,
        storeName: martName,
        storeAddress: address,
        latitude: latitude ?? undefined,
        longitude: longitude ?? undefined,
        geocodeStatus: latitude !== null && longitude !== null ? "성공" : "미시도",
        mapIncluded: true,
        operatingStatus: "영업 중",
        itemCount: 0,
        completedCount: 0,
        surveyDate: today(),
        status: "미시작",
        updatedAt: current,
      });
    }
    const store = storeMap.get(key)!;
    store.itemCount += 1;
    if (!regionMap.has(resolved.name)) {
      regionMap.set(resolved.name, { name: resolved.name, department: resolved.department, city: resolved.city, areaSummary: "", updatedAt: current });
    }
    items.push({
      id: uid("item"),
      itemNo: pick(row, "itemNo") || `용${String(index + 1).padStart(4, "0")}`,
      sourceOrder: items.length,
      region: resolved.name,
      department: resolved.department,
      city: resolved.city,
      storeId: store.id,
      storeName: store.storeName,
      storeAddress: store.storeAddress,
      detailAddress,
      companyName: pick(row, "companyName"),
      companyTel: pick(row, "companyTel"),
      companyManager: pick(row, "companyManager"),
      martName,
      martTel: pick(row, "martTel"),
      barcode,
      productImageUrl: pick(row, "productImageUrl"),
      productName,
      spec: pick(row, "spec"),
      basePrice: numberOrNull(pick(row, "basePrice")),
      surveyDate: store.surveyDate,
      normalDisplay: "",
      specMatch: "",
      barcodeMatch: "",
      normalPrice: null,
      hasDiscount: false,
      discountPrice: null,
      discountStartDate: "",
      discountEndDate: "",
      discountType: "",
      discountOral: false,
      discountPeriodMode: "",
      priceJudgment: "",
      abnormalDisplay: "",
      barcodeRegistered: "",
      abnormalStatus: "",
      posChecked: "",
      posPrice: null,
      memo: "",
      status: "미조사",
      updatedAt: current,
    });
  });

  // 시군구는 입력 파일에서 확인된 값만 중복 제거하여 카드의 부제목으로 보낸다.
  for (const [city, region] of regionMap) {
    region.areaSummary = Array.from(districtMap.get(city) || []).sort((a, b) => a.localeCompare(b, "ko")).join(", ");
  }

  return { regions: sortRegions(Array.from(regionMap.values())), stores: Array.from(storeMap.values()), items };
}

export function mergeContacts(items: SurveyItem[], contactFileRows: Row[]) {
  const contacts = new Map<string, { tel: string; manager: string; address: string; martName: string; martTel: string; region: string; city: string; department: string }>();
  const byItemNo = new Map<string, { tel: string; manager: string; address: string; martName: string; martTel: string; region: string; city: string; department: string }>();
  contactFileRows.forEach((row) => {
    const company = pick(row, "companyName");
    const contactCity = pick(row, "city");
    const contact = {
      tel: pick(row, "companyTel"),
      manager: pick(row, "companyManager"),
      address: pick(row, "address"),
      martName: pick(row, "martName"),
      martTel: pick(row, "martTel"),
      region: contactCity,
      city: contactCity,
      department: pick(row, "department"),
    };
    if (company) contacts.set(company, contact);
    const itemNo = pick(row, "itemNo");
    if (itemNo) byItemNo.set(itemNo, contact);
  });

  return items.map((item) => {
    const contact = byItemNo.get(item.itemNo) ?? contacts.get(item.companyName);
    return contact
      ? {
          ...item,
          // 연락처 보완 파일이 원본 조사표의 지역을 덮어쓰지 않도록 한다.
          region: item.region,
          city: item.city,
          department: item.department,
          companyTel: item.companyTel || contact.tel,
          companyManager: item.companyManager || contact.manager,
          martTel: item.martTel || contact.martTel,
          storeAddress: item.storeAddress || contact.address,
          detailAddress: item.detailAddress || contact.address,
          martName: item.martName || contact.martName,
          storeName: item.storeName || contact.martName,
        }
      : item;
  });
}

export async function parseContactRows(source: Blob | ArrayBuffer): Promise<Row[]> {
  const workbook = await readWorkbook(source);
  const sheetNames = dataSheetNames(workbook);
  return sheetNames.flatMap((name) => rowsFromSheet(workbook.Sheets[name], name));
}

export function rebuildStoresAndRegions(items: SurveyItem[], sourceRegions: Region[] = []): { regions: Region[]; stores: SurveyStore[]; items: SurveyItem[] } {
  const current = now();
  const regionMap = new Map<string, Region>();
  const storeMap = new Map<string, SurveyStore>();
  const sourceRegionByName = new Map(sourceRegions.map((region) => [region.name, region]));
  const rebuiltItems = items.map((item) => {
    // 이미 엑셀에서 확정된 광역 이름을 변경하지 않는다.
    const regionName = item.region || item.city || "미지정";
    const department = undefined;
    const city = item.city || regionName;
    if (!regionMap.has(regionName)) {
      regionMap.set(regionName, { name: regionName, department, city, areaSummary: sourceRegionByName.get(regionName)?.areaSummary || "", updatedAt: current });
    }
    const key = storeKey(regionName, item.storeName || item.martName, item.storeAddress);
    if (!storeMap.has(key)) {
      storeMap.set(key, {
        id: uid("store"),
        region: regionName,
        department,
        city,
        storeName: item.storeName || item.martName || "미상 방문지",
        storeAddress: item.storeAddress,
        operatingStatus: "영업 중",
        itemCount: 0,
        completedCount: 0,
        surveyDate: today(),
        status: "미시작",
        updatedAt: current,
      });
    }
    const store = storeMap.get(key)!;
    store.itemCount += 1;
    return { ...item, region: regionName, department, city, storeId: store.id, storeName: store.storeName, storeAddress: store.storeAddress, updatedAt: current };
  });

  return { regions: sortRegions(Array.from(regionMap.values())), stores: Array.from(storeMap.values()), items: rebuiltItems };
}

function sortRegions(regions: Region[]) {
  // 업로드 엑셀에서 광역이 처음 등장한 순서를 유지한다.
  return regions;
}

function rowsFromSheet(sheet: XLSX.WorkSheet, sheetName = ""): Row[] {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
  const headerIndex = matrix.findIndex((row) => {
    const joined = row.map(clean).join(" ");
    return joined.includes("순번") && (joined.includes("물품명") || joined.includes("품목명") || joined.includes("상품명"));
  });
  if (headerIndex === -1) return XLSX.utils.sheet_to_json<Row>(sheet, { defval: "" }).map((row) => ({ ...row, __sheetName: sheetName }));
  const headers = matrix[headerIndex].map((header, index) => clean(header) || `빈열${index}`);
  return matrix.slice(headerIndex + 1).map((row) => ({ ...Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])), __sheetName: sheetName }));
}
