import * as XLSX from "xlsx";
import { now, today, uid } from "./db";
import type { Region, SurveyItem, SurveyStore } from "./types";

type Row = Record<string, unknown>;

type RegionRule = {
  name: string;
  cityNames: string[];
  districts: string[];
  department: string;
  areaSummary?: string;
};

const REGION_RULES: RegionRule[] = [
  { name: "서울 1권역", cityNames: ["서울"], districts: ["마포", "서대문", "용산", "은평", "종로", "중구"], department: "조사1부", areaSummary: "마포, 서대문, 용산, 은평, 종로, 중구" },
  { name: "서울 2권역", cityNames: ["서울"], districts: ["강북", "광진", "노원", "도봉", "동대문", "성동", "성북", "중랑"], department: "조사2부", areaSummary: "강북, 광진, 노원, 도봉, 동대문, 성동, 성북, 중랑" },
  { name: "서울 3권역", cityNames: ["서울"], districts: ["강서", "관악", "구로", "금천", "동작", "양천", "영등포"], department: "회원부", areaSummary: "강서, 관악, 구로, 금천, 동작, 양천, 영등포" },
  { name: "서울 4권역", cityNames: ["서울"], districts: ["강남", "강동", "서초", "송파"], department: "조사2부 / 전산사업부", areaSummary: "강남, 강동, 서초, 송파" },
  { name: "경기 1권역", cityNames: ["경기"], districts: ["고양", "김포", "파주", "양주", "연천", "의정부", "포천"], department: "기획조사부", areaSummary: "고양, 김포, 파주, 양주, 연천, 의정부, 포천" },
  { name: "경기 2권역", cityNames: ["경기"], districts: ["구리", "남양주", "동두천"], department: "조사2부", areaSummary: "구리, 남양주, 동두천" },
  { name: "경기 3권역", cityNames: ["경기"], districts: ["광명", "부천", "시흥", "안산", "화성"], department: "경영지원부", areaSummary: "광명, 부천, 시흥, 안산, 화성" },
  { name: "경기 4권역", cityNames: ["경기"], districts: ["과천", "군포", "수원", "안성", "안양", "의왕", "평택"], department: "원가조사부", areaSummary: "과천, 군포, 수원, 안성, 안양, 의왕, 평택" },
  { name: "경기 5권역", cityNames: ["경기"], districts: ["광주", "성남", "양평", "여주", "용인", "이천", "하남"], department: "조사1부", areaSummary: "광주, 성남, 양평, 여주, 용인, 이천, 하남" },
  { name: "인천", cityNames: ["인천"], districts: [], department: "회원부 / 기획조사부" },
  { name: "울산", cityNames: ["울산"], districts: [], department: "조사2부" },
  { name: "강원", cityNames: ["강원"], districts: [], department: "조사2부" },
  { name: "제주", cityNames: ["제주"], districts: [], department: "조사2부" },
  { name: "대전", cityNames: ["대전"], districts: [], department: "지회 충청권" },
  { name: "세종", cityNames: ["세종"], districts: [], department: "지회 충청권" },
  { name: "충남", cityNames: ["충남", "충청남도"], districts: [], department: "지회 충청권" },
  { name: "충북", cityNames: ["충북", "충청북도"], districts: [], department: "지회 충청권" },
  { name: "대구", cityNames: ["대구"], districts: [], department: "지회 경상권" },
  { name: "경북", cityNames: ["경북", "경상북도"], districts: [], department: "지회 경상권" },
  { name: "광주", cityNames: ["광주"], districts: [], department: "지회 전라권" },
  { name: "전남", cityNames: ["전남", "전라남도"], districts: [], department: "지회 전라권" },
  { name: "전북", cityNames: ["전북", "전라북도", "전북특별자치도"], districts: [], department: "지회 전라권" },
  { name: "부산", cityNames: ["부산"], districts: [], department: "지회 경남권" },
  { name: "경남", cityNames: ["경남", "경상남도"], districts: [], department: "지회 경남권" },
];

const SHEET_REGION_ALIASES = new Map([
  ["서울1", "서울 1권역"],
  ["서울2", "서울 2권역"],
  ["서울3", "서울 3권역"],
  ["서울4", "서울 4권역"],
  ["경기1", "경기 1권역"],
  ["경기2", "경기 2권역"],
  ["경기3", "경기 3권역"],
  ["경기4", "경기 4권역"],
  ["경기5", "경기 5권역"],
]);

const aliases: Record<string, string[]> = {
  itemNo: ["순번", "번호", "품목번호", "연번", "no"],
  companyName: ["업체명", "회사명", "제조회사", "제조사", "공급업체"],
  companyTel: ["업체연락처", "담당자연락처", "담당자 연락처", "연락처", "전화번호", "업체전화", "TEL"],
  companyManager: ["담당자", "담당", "업체담당자"],
  martName: ["마트명", "판매처", "조사처", "매장명", "방문지", "마트"],
  martTel: ["마트번호", "마트 연락처", "마트연락처", "판매처번호", "판매처연락처", "매장번호", "매장연락처"],
  barcode: ["바코드", "barcode", "상품코드"],
  productName: ["물품명", "품목명", "상품명", "제품명"],
  spec: ["규격", "용량", "단위"],
  basePrice: ["기준가격", "기준가", "가격", "기준단가"],
  region: ["권역", "담당권역"],
  city: ["광역", "시도"],
  district: ["지역", "시군구", "구군"],
  department: ["담당부서", "담당", "부서", "지회", "담당지회"],
  address: ["주소", "마트주소", "판매처주소", "매장주소"],
  finalAddress: ["최종주소", "최종 주소", "지도주소", "검색주소"],
  detailAddress: ["주소", "상세주소", "세부주소"],
  latitude: ["위도", "latitude", "lat", "y좌표", "y"],
  longitude: ["경도", "longitude", "lng", "lon", "x좌표", "x"],
};

const clean = (value: unknown) => String(value ?? "").trim();
const compact = (value: string) => value.toLowerCase().replace(/[\s_\-()[\]{}]/g, "");
const normalizePlace = (value: string) =>
  value
    .replace(/\(.+?\)/g, "")
    .replace(/특별자치도|특별자치시|광역시|특별시|자치시|자치도|시|군|구/g, "")
    .replace(/\s/g, "");

function pick(row: Row, key: keyof typeof aliases) {
  const headers = Object.keys(row);
  const found = headers.find((header) => aliases[key].some((alias) => compact(header).includes(compact(alias))));
  return found ? clean(row[found]) : "";
}

function numberOrNull(value: string) {
  const parsed = Number(value.replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function ruleByName(name: string) {
  return REGION_RULES.find((rule) => rule.name === name);
}

function resolveSurveyRegion(city: string, district: string, fallback = "미지정") {
  const alias = SHEET_REGION_ALIASES.get(fallback.replace(/\s/g, ""));
  if (alias) {
    const rule = ruleByName(alias);
    return { name: alias, department: rule?.department, city: rule?.cityNames[0], areaSummary: rule?.areaSummary };
  }

  const normalizedCity = normalizePlace(city);
  const normalizedDistrict = normalizePlace(district);
  const rule = REGION_RULES.find((candidate) => {
    const cityMatched = candidate.cityNames.some((name) => normalizedCity.includes(normalizePlace(name)) || normalizePlace(name).includes(normalizedCity));
    if (!cityMatched) return false;
    if (candidate.districts.length === 0) return true;
    return candidate.districts.some((name) => normalizedDistrict.includes(normalizePlace(name)));
  });

  if (rule) return { name: rule.name, department: rule.department, city: rule.cityNames[0], areaSummary: rule.areaSummary };
  const direct = ruleByName(fallback);
  if (direct) return { name: direct.name, department: direct.department, city: direct.cityNames[0], areaSummary: direct.areaSummary };
  return { name: fallback || "미지정", department: undefined, city, areaSummary: undefined };
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
  const items: SurveyItem[] = [];
  const current = now();

  sheetNames.flatMap((name) => rowsFromSheet(workbook.Sheets[name], name).map((row) => ({ row, sheetName: name }))).forEach(({ row, sheetName }, index) => {
    const productName = pick(row, "productName");
    const barcode = pick(row, "barcode");
    if (!productName && !barcode) return;
    const resolved = resolveSurveyRegion(pick(row, "city"), pick(row, "district"), pick(row, "region") || sheetName);
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
      regionMap.set(resolved.name, { name: resolved.name, department: resolved.department, city: resolved.city, areaSummary: resolved.areaSummary, updatedAt: current });
    }
    items.push({
      id: uid("item"),
      itemNo: pick(row, "itemNo") || `용${String(index + 1).padStart(4, "0")}`,
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
      productName,
      spec: pick(row, "spec"),
      basePrice: numberOrNull(pick(row, "basePrice")),
      surveyDate: store.surveyDate,
      normalDisplay: "",
      specMatch: "",
      barcodeMatch: "",
      normalPrice: null,
      hasDiscount: null,
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

  return { regions: sortRegions(Array.from(regionMap.values())), stores: Array.from(storeMap.values()), items };
}

export function mergeContacts(items: SurveyItem[], contactFileRows: Row[]) {
  const contacts = new Map<string, { tel: string; manager: string; address: string; martName: string; martTel: string; region: string; city: string; department: string }>();
  const byItemNo = new Map<string, { tel: string; manager: string; address: string; martName: string; martTel: string; region: string; city: string; department: string }>();
  contactFileRows.forEach((row) => {
    const company = pick(row, "companyName");
    const resolved = resolveSurveyRegion(pick(row, "city"), pick(row, "district"), clean(row.__sheetName));
    const contact = {
      tel: pick(row, "companyTel"),
      manager: pick(row, "companyManager"),
      address: pick(row, "address"),
      martName: pick(row, "martName"),
      martTel: pick(row, "martTel"),
      region: resolved.name,
      city: resolved.city || pick(row, "city"),
      department: resolved.department || pick(row, "department"),
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
          region: contact.region || item.region,
          city: contact.city || item.city,
          department: contact.department || item.department,
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

export function rebuildStoresAndRegions(items: SurveyItem[]): { regions: Region[]; stores: SurveyStore[]; items: SurveyItem[] } {
  const current = now();
  const regionMap = new Map<string, Region>();
  const storeMap = new Map<string, SurveyStore>();
  const rebuiltItems = items.map((item) => {
    const resolved = resolveSurveyRegion(item.city || "", "", item.region || "미지정");
    const regionName = resolved.name;
    const department = item.department || resolved.department;
    const city = item.city || resolved.city;
    if (!regionMap.has(regionName)) {
      regionMap.set(regionName, { name: regionName, department, city, areaSummary: resolved.areaSummary, updatedAt: current });
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
  const order = REGION_RULES.map((rule) => rule.name);
  return regions.sort((a, b) => {
    const ai = order.indexOf(a.name);
    const bi = order.indexOf(b.name);
    if (ai !== -1 || bi !== -1) return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
    return a.name.localeCompare(b.name, "ko");
  });
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
