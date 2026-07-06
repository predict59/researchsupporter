import { Camera, CheckCircle2, ChevronDown, ChevronUp, Download, Menu, MoreVertical, Phone, SlidersHorizontal, Search, Upload, X } from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { clearAllData, deletePhoto, getBarcodeIndex, getItems, getPhotos, getPhotosByRegion, getPhotosByStore, getRegions, getSettings, getStores, importAllData, importRegionData, now, putItem, putPhoto, putStore, saveBarcodeIndex, saveParsedData, saveSettings, today, uid } from "./db";
import { extractBarcodeImages } from "./barcodeImages";
import { parseContactRows, parseSurveyWorkbook, mergeContacts, rebuildStoresAndRegions } from "./excel";
import { dataUrlToBlob, exportBackup, exportRegionExcel, exportRegionZip } from "./exporters";
import { mapSearchAddress, requiredPhotoLabels, summarize } from "./logic";
import type { AppSettings, BackupPayload, PhotoType, Region, RegionStats, StoreOperatingStatus, SurveyItem, SurveyPhoto, SurveyStore } from "./types";

type View = "upload" | "regions" | "assignment" | "workspace" | "store" | "items" | "item" | "backup" | "validation";
type Filter = "전체" | "미완료" | "미조사" | "조사중" | "완료" | "사진누락" | "미진열";
type StoreSort = "이름 순" | "품목 많은 순" | "미완료 많은 순" | "거리 순";
type WorkspaceMode = "list" | "map";
type ConfirmState = {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
  plain?: boolean;
};
type AppHistoryState = {
  app: "research-supporter";
  view: View;
  currentRegion?: string;
  selectedStoreId?: string;
  selectedItemId?: string;
  workspaceMode?: WorkspaceMode;
  barcodeModalItemId?: string;
};

const mapLinks = (address: string) => [
  ["구글", `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(mapSearchAddress(address))}`],
  ["네이버", `https://map.naver.com/p/search/${encodeURIComponent(mapSearchAddress(address))}`],
  ["카카오", `https://map.kakao.com/link/search/${encodeURIComponent(mapSearchAddress(address))}`],
];
const delay = (ms: number) => new Promise((resolve) => window.setTimeout(resolve, ms));

const emptyStats: RegionStats = { total: 0, completed: 0, inProgress: 0, notStarted: 0, photoMissing: 0 };
const num = (value: string) => {
  const digits = value.replace(/\D/g, "");
  return digits === "" ? null : Number(digits);
};
const PHOTO_MAX_EDGE = 1280;
const PHOTO_TARGET_BYTES = 950 * 1024;
const PHOTO_MIN_EDGE = 760;
const PHOTO_QUALITY_STEPS = [0.72, 0.64, 0.56, 0.48, 0.4, 0.32];
const PRICE_DIFF_WARN_PERCENT = 30;
const TARGET_MAP_URL = "https://www.google.com/maps/d/u/1/edit?mid=1ej99Lo6WS4GROBCQPr0a66MhQR_vXuM&usp=sharing";
type BarcodeImageIndex = Record<string, string>;
type PriceCandidate = { value: number; score: number; source: "comma" | "plain" };
type PriceOcrWorker = Awaited<ReturnType<typeof import("tesseract.js")["createWorker"]>>;
const PRICE_KEYWORDS = /원|가격|정상|판매|할인|행사|특가|세일|SALE|sale|올리브|카드|멤버십|회원|쿠폰/;
const PRICE_MAX_VALUE = 999999;
type GeocodeResult = { latitude: number; longitude: number; displayName?: string };
const appendMemoText = (memo: string, text: string) => {
  const parts = memo.split("/").map((part) => part.trim()).filter(Boolean);
  if (parts.includes(text)) return memo;
  return parts.length ? `${parts.join(" / ")} / ${text}` : text;
};
const removeMemoTexts = (memo: string, texts: string[]) =>
  memo.split("/").map((part) => part.trim()).filter((part) => part && !texts.includes(part)).join(" / ");
const STORE_STATUS_MEMOS = ["판매처 폐점", "임시휴업"];
const POS_MEMOS = ["POS 조회", "POS 조회 불가", "POS 확인"];
const periodTypeFromDates = (start: string, end: string) => {
  if (!start || !end) return "";
  const startTime = new Date(`${start}T00:00:00`).getTime();
  const endTime = new Date(`${end}T00:00:00`).getTime();
  if (!Number.isFinite(startTime) || !Number.isFinite(endTime) || endTime < startTime) return "";
  const days = Math.floor((endTime - startTime) / 86400000) + 1;
  return days <= 31 ? "①" : "②";
};
type BarcodeDetectorConstructor = new (options?: { formats?: string[] }) => {
  detect: (source: ImageBitmapSource) => Promise<Array<{ rawValue: string }>>;
};
const barcodeFormats = ["ean_13", "ean_8", "code_128", "code_39", "code_93", "upc_a", "upc_e", "itf"];
const onlyDigits = (value: string) => value.replace(/\D/g, "");
const formatBytes = (value?: number) => {
  if (!value) return "-";
  const units = ["B", "KB", "MB", "GB"];
  let size = value;
  let unit = 0;
  while (size >= 1024 && unit < units.length - 1) {
    size /= 1024;
    unit += 1;
  }
  return `${size.toFixed(unit === 0 ? 0 : 1)} ${units[unit]}`;
};
const availableStorageBytes = (estimate?: StorageEstimate) => Math.max(0, (estimate?.quota ?? 0) - (estimate?.usage ?? 0));
const distanceKm = (from: { latitude: number; longitude: number }, to: { latitude: number; longitude: number }) => {
  const rad = (value: number) => value * Math.PI / 180;
  const earth = 6371;
  const dLat = rad(to.latitude - from.latitude);
  const dLon = rad(to.longitude - from.longitude);
  const lat1 = rad(from.latitude);
  const lat2 = rad(to.latitude);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) ** 2;
  return earth * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
};
const formatDistance = (km?: number) => {
  if (km === undefined) return "";
  return km < 1 ? `${Math.round(km * 1000)}m` : `${km.toFixed(km < 10 ? 1 : 0)}km`;
};
const barcodeImageSrc = (item: SurveyItem, index: BarcodeImageIndex) => {
  const path = index[item.itemNo];
  if (!path) return "";
  return path.startsWith("data:") ? path : `${import.meta.env.BASE_URL}${path}`;
};
let priceOcrWorkerPromise: Promise<PriceOcrWorker> | null = null;

function barcodeScanRegions(width: number, height: number) {
  const regions = [{ x: 0, y: 0, width, height }];
  const addGrid = (cols: number, rows: number, overlap = 0.18) => {
    const cellW = width / cols;
    const cellH = height / rows;
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const x = Math.max(0, col * cellW - cellW * overlap);
        const y = Math.max(0, row * cellH - cellH * overlap);
        const right = Math.min(width, (col + 1) * cellW + cellW * overlap);
        const bottom = Math.min(height, (row + 1) * cellH + cellH * overlap);
        regions.push({ x, y, width: right - x, height: bottom - y });
      }
    }
  };
  addGrid(2, 2);
  addGrid(3, 3);
  addGrid(4, 4, 0.22);
  regions.push(
    { x: 0, y: 0, width, height: height / 2 },
    { x: 0, y: height / 2, width, height: height / 2 },
    { x: 0, y: 0, width: width / 2, height },
    { x: width / 2, y: 0, width: width / 2, height },
    { x: 0, y: height * 0.25, width, height: height * 0.5 },
    { x: width * 0.25, y: 0, width: width * 0.5, height },
    { x: width * 0.15, y: height * 0.15, width: width * 0.7, height: height * 0.7 },
  );
  return regions;
}

function cropToCanvas(source: ImageBitmap, region: { x: number; y: number; width: number; height: number }, enhance = false) {
  const canvas = document.createElement("canvas");
  const scale = Math.min(4, Math.max(1.2, 2200 / Math.max(region.width, region.height)));
  canvas.width = Math.round(region.width * scale);
  canvas.height = Math.round(region.height * scale);
  const context = canvas.getContext("2d");
  if (!context) return undefined;
  context.imageSmoothingEnabled = false;
  context.drawImage(source, region.x, region.y, region.width, region.height, 0, 0, canvas.width, canvas.height);
  if (enhance) {
    const image = context.getImageData(0, 0, canvas.width, canvas.height);
    const data = image.data;
    for (let index = 0; index < data.length; index += 4) {
      const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
      const boosted = gray > 150 ? 255 : gray < 95 ? 0 : gray;
      data[index] = boosted;
      data[index + 1] = boosted;
      data[index + 2] = boosted;
    }
    context.putImageData(image, 0, 0);
  }
  return canvas;
}

async function detectBarcodeFromFile(file: File) {
  const detectorClass = (window as typeof window & { BarcodeDetector?: BarcodeDetectorConstructor }).BarcodeDetector;
  if (!detectorClass) return { supported: false, values: [] as string[] };
  const bitmap = await createImageBitmap(file);
  try {
    const detector = new detectorClass({ formats: barcodeFormats });
    const values = new Set<string>();
    for (const region of barcodeScanRegions(bitmap.width, bitmap.height)) {
      const sources: ImageBitmapSource[] = [];
      const isFull = region.x === 0 && region.y === 0 && region.width === bitmap.width && region.height === bitmap.height;
      if (isFull) sources.push(bitmap);
      const cropped = cropToCanvas(bitmap, region);
      const enhanced = cropToCanvas(bitmap, region, true);
      if (cropped) sources.push(cropped);
      if (enhanced) sources.push(enhanced);
      for (const source of sources) {
        const results = await detector.detect(source);
        results.map((result) => result.rawValue).filter(Boolean).forEach((value) => values.add(value));
        if (values.size > 0) break;
      }
      if (values.size > 0) break;
      await new Promise((resolve) => window.setTimeout(resolve, 0));
    }
    return { supported: true, values: Array.from(values) };
  } finally {
    bitmap.close();
  }
}

const addPriceCandidate = (bucket: Map<number, PriceCandidate>, text: string, raw: string, index: number, source: PriceCandidate["source"]) => {
  const digits = raw.replace(/\D/g, "");
  if (digits.length < 3 || digits.length > 6) return;
  const value = Number(digits);
  if (!Number.isFinite(value) || value < 100 || value > PRICE_MAX_VALUE) return;

  const context = text.slice(Math.max(0, index - 24), Math.min(text.length, index + raw.length + 24));
  let score = source === "comma" ? 60 : 24;
  if (PRICE_KEYWORDS.test(context)) score += 28;
  if (/[원￦₩]/.test(context)) score += 18;
  if (/할인|행사|특가|세일|SALE|sale/.test(context)) score += 12;
  if (/%|g|kg|ml|L|개입|입|매|번|호|월|일/.test(context)) score -= 12;
  if (value % 10 !== 0) score -= 6;
  if (value >= 1000 && value <= 300000) score += 8;

  const existing = bucket.get(value);
  if (!existing || existing.score < score) bucket.set(value, { value, score, source });
};

function extractPriceCandidates(text: string) {
  const normalized = text
    .replace(/[０-９]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0))
    .replace(/[，]/g, ",")
    .replace(/(\d)\s*,\s*(\d{3})/g, "$1,$2");
  const bucket = new Map<number, PriceCandidate>();

  for (const match of normalized.matchAll(/\d{1,3}(?:,\d{3})+/g)) {
    addPriceCandidate(bucket, normalized, match[0], match.index ?? 0, "comma");
  }
  for (const match of normalized.matchAll(/(?:^|[^\d,])(\d{3,6})(?![\d,])/g)) {
    const raw = match[1];
    addPriceCandidate(bucket, normalized, raw, (match.index ?? 0) + match[0].indexOf(raw), "plain");
  }

  return Array.from(bucket.values())
    .filter((candidate) => candidate.score > 18)
    .sort((a, b) => b.score - a.score || b.value - a.value)
    .slice(0, 4);
}

function createPriceOcrCanvas(bitmap: ImageBitmap, mode: "contrast" | "threshold", crop?: { x: number; y: number; width: number; height: number }) {
  const source = crop ?? { x: 0, y: 0, width: bitmap.width, height: bitmap.height };
  const canvas = document.createElement("canvas");
  const scale = Math.min(2.4, Math.max(1.4, 2200 / Math.max(source.width, source.height)));
  canvas.width = Math.round(source.width * scale);
  canvas.height = Math.round(source.height * scale);
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return undefined;
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  context.drawImage(bitmap, source.x, source.y, source.width, source.height, 0, 0, canvas.width, canvas.height);

  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  const data = image.data;
  for (let index = 0; index < data.length; index += 4) {
    const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
    const adjusted = mode === "threshold"
      ? gray > 142 ? 255 : 0
      : Math.max(0, Math.min(255, (gray - 128) * 1.65 + 128));
    data[index] = adjusted;
    data[index + 1] = adjusted;
    data[index + 2] = adjusted;
  }
  context.putImageData(image, 0, 0);
  return canvas;
}

async function createPriceOcrSources(blob: Blob) {
  const bitmap = await createImageBitmap(blob);
  try {
    const sources: HTMLCanvasElement[] = [];
    const full = { x: 0, y: 0, width: bitmap.width, height: bitmap.height };
    const center = { x: bitmap.width * 0.08, y: bitmap.height * 0.08, width: bitmap.width * 0.84, height: bitmap.height * 0.84 };
    const top = { x: 0, y: 0, width: bitmap.width, height: bitmap.height * 0.7 };
    [full, center, top].forEach((crop, index) => {
      const contrast = createPriceOcrCanvas(bitmap, "contrast", crop);
      if (contrast) sources.push(contrast);
      if (index < 2) {
        const threshold = createPriceOcrCanvas(bitmap, "threshold", crop);
        if (threshold) sources.push(threshold);
      }
    });
    return sources;
  } finally {
    bitmap.close();
  }
}

async function getPriceOcrWorker() {
  const tesseract = await import("tesseract.js");
  if (!priceOcrWorkerPromise) {
    priceOcrWorkerPromise = (async () => {
      const worker = await tesseract.createWorker("eng", tesseract.OEM.LSTM_ONLY, {
        logger: () => undefined,
      });
      await worker.setParameters({
        tessedit_pageseg_mode: tesseract.PSM.SPARSE_TEXT,
        tessedit_char_whitelist: "0123456789,.",
        preserve_interword_spaces: "1",
        user_defined_dpi: "300",
      });
      return worker;
    })();
  }
  return priceOcrWorkerPromise;
}

async function detectPriceCandidatesFromBlob(blob: Blob) {
  const worker = await getPriceOcrWorker();
  try {
    const sources = await createPriceOcrSources(blob);
    const merged = new Map<number, PriceCandidate>();
    for (const source of sources) {
      const result = await worker.recognize(source);
      for (const candidate of extractPriceCandidates(result.data.text)) {
        const old = merged.get(candidate.value);
        const boosted = { ...candidate, score: candidate.score + (old ? 12 : 0) };
        if (!old || old.score < boosted.score) merged.set(candidate.value, boosted);
      }
      if (merged.size >= 4) {
        const bestScore = Math.max(...Array.from(merged.values()).map((candidate) => candidate.score));
        if (bestScore >= 90) break;
      }
      await new Promise((resolve) => window.setTimeout(resolve, 0));
    }
    return Array.from(merged.values()).sort((a, b) => b.score - a.score || b.value - a.value).slice(0, 4);
  } catch (error) {
    priceOcrWorkerPromise = null;
    throw error;
  }
}

async function geocodeAddress(address: string): Promise<GeocodeResult | null> {
  const query = mapSearchAddress(address);
  if (!query) return null;
  const url = new URL("https://nominatim.openstreetmap.org/search");
  url.searchParams.set("format", "jsonv2");
  url.searchParams.set("q", query);
  url.searchParams.set("limit", "1");
  url.searchParams.set("countrycodes", "kr");
  url.searchParams.set("addressdetails", "0");
  const response = await fetch(url.toString(), {
    headers: { Accept: "application/json", "Accept-Language": "ko" },
  });
  if (!response.ok) return null;
  const results = await response.json() as Array<{ lat?: string; lon?: string; display_name?: string }>;
  const first = results[0];
  if (!first?.lat || !first.lon) return null;
  const latitude = Number(first.lat);
  const longitude = Number(first.lon);
  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) return null;
  return { latitude, longitude, displayName: first.display_name };
}

async function resizePhoto(file: File) {
  if (!file.type.startsWith("image/")) return { blob: file, mimeType: file.type || "application/octet-stream", originalSize: file.size, resizedSize: file.size };
  const bitmap = await createImageBitmap(file);
  try {
    const sourceEdge = Math.max(bitmap.width, bitmap.height);
    const edgeSteps = [PHOTO_MAX_EDGE, 1150, 1024, 900, PHOTO_MIN_EDGE].filter((edge, index, array) => edge <= sourceEdge && array.indexOf(edge) === index);
    if (!edgeSteps.length) edgeSteps.push(sourceEdge);
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    if (!context) return { blob: file, mimeType: file.type || "application/octet-stream", originalSize: file.size, resizedSize: file.size };
    let best: Blob | null = null;
    for (const maxEdge of edgeSteps) {
      const scale = Math.min(1, maxEdge / sourceEdge);
      const width = Math.max(1, Math.round(bitmap.width * scale));
      const height = Math.max(1, Math.round(bitmap.height * scale));
      canvas.width = width;
      canvas.height = height;
      context.clearRect(0, 0, width, height);
      context.drawImage(bitmap, 0, 0, width, height);
      for (const quality of PHOTO_QUALITY_STEPS) {
        const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/jpeg", quality));
        if (!blob) continue;
        if (!best || blob.size < best.size) best = blob;
        if (blob.size <= PHOTO_TARGET_BYTES) return { blob, mimeType: "image/jpeg", originalSize: file.size, resizedSize: blob.size };
      }
    }
    const output = best && best.size < file.size ? best : file;
    return { blob: output, mimeType: output.type || file.type || "image/jpeg", originalSize: file.size, resizedSize: output.size };
  } finally {
    bitmap.close();
  }
}

function App() {
  const [view, setView] = useState<View>("upload");
  const topbarRef = useRef<HTMLElement | null>(null);
  const [isBooting, setIsBooting] = useState(true);
  const [settings, setSettingsState] = useState<AppSettings>({ defaultSurveyDate: today() });
  const [regions, setRegions] = useState<Region[]>([]);
  const [stores, setStores] = useState<SurveyStore[]>([]);
  const [items, setItems] = useState<SurveyItem[]>([]);
  const [photos, setPhotos] = useState<SurveyPhoto[]>([]);
  const [selectedStoreId, setSelectedStoreId] = useState("");
  const [selectedItemId, setSelectedItemId] = useState("");
  const [mapFocusStoreId, setMapFocusStoreId] = useState("");
  const [regionQuery, setRegionQuery] = useState("");
  const [storeQuery, setStoreQuery] = useState("");
  const [itemQuery, setItemQuery] = useState("");
  const [filter, setFilter] = useState<Filter>("전체");
  const [barcodeIndex, setBarcodeIndex] = useState<BarcodeImageIndex>({});
  const [barcodeModalItemId, setBarcodeModalItemId] = useState("");
  const [barcodeReturnItemId, setBarcodeReturnItemId] = useState("");
  const [surveyFile, setSurveyFile] = useState<File | null>(null);
  const [contactFile, setContactFile] = useState<File | null>(null);
  const [barcodeFile, setBarcodeFile] = useState<File | null>(null);
  const [uploadMessage, setUploadMessage] = useState("");
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [storeSort, setStoreSort] = useState<StoreSort>("이름 순");
  const [workspaceMode, setWorkspaceMode] = useState<WorkspaceMode>("list");
  const [workspaceToolsOpen, setWorkspaceToolsOpen] = useState(false);
  const [itemToolsOpen, setItemToolsOpen] = useState(false);
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [contactStoreId, setContactStoreId] = useState("");
  const [storageOpen, setStorageOpen] = useState(false);
  const [storageEstimate, setStorageEstimate] = useState<StorageEstimate | undefined>();
  const [frontPhotoPickerOpen, setFrontPhotoPickerOpen] = useState(false);
  const [userLocation, setUserLocation] = useState<{ latitude: number; longitude: number } | null>(null);
  const [locationFocusTick, setLocationFocusTick] = useState(0);
  const [geocodeMessage, setGeocodeMessage] = useState("");
  const [geocoding, setGeocoding] = useState(false);
  const [photosReady, setPhotosReady] = useState(false);
  const [confirmState, setConfirmState] = useState<ConfirmState | null>(null);
  const [storeStatusDraft, setStoreStatusDraft] = useState<StoreOperatingStatus | "">("");
  const [storeStatusMessage, setStoreStatusMessage] = useState("");
  const confirmResolver = useRef<((value: boolean) => void) | null>(null);
  const locatePromiseRef = useRef<Promise<{ latitude: number; longitude: number } | null> | null>(null);
  const initialLocationRequested = useRef(false);
  const historyReadyRef = useRef(false);
  const restoringHistoryRef = useRef(false);
  const lastHistoryViewRef = useRef<View>("upload");
  const viewRef = useRef<View>("upload");
  const lastExitBackRef = useRef(0);
  const exitToastTimerRef = useRef<number | undefined>();
  const allowBrowserExitRef = useRef(false);
  const [appBackToast, setAppBackToast] = useState("");

  const currentRegion = settings.currentRegion;
  const regionItems = useMemo(() => items.filter((item) => item.region === currentRegion), [items, currentRegion]);
  const regionStores = useMemo(() => stores.filter((store) => store.region === currentRegion), [stores, currentRegion]);
  const photosByStore = useMemo(() => {
    const map = new Map<string, SurveyPhoto[]>();
    photos.forEach((photo) => map.set(photo.storeId, [...(map.get(photo.storeId) ?? []), photo]));
    return map;
  }, [photos]);
  const regionItemsByStore = useMemo(() => {
    const map = new Map<string, SurveyItem[]>();
    regionItems.forEach((item) => map.set(item.storeId, [...(map.get(item.storeId) ?? []), item]));
    return map;
  }, [regionItems]);
  const regionStatsByStore = useMemo(() => {
    const map = new Map<string, RegionStats>();
    regionStores.forEach((store) => map.set(store.id, summarize(regionItemsByStore.get(store.id) ?? [], photosByStore.get(store.id) ?? [])));
    return map;
  }, [regionStores, regionItemsByStore, photosByStore]);
  const sortedRegionStores = useMemo(() => {
    return [...regionStores].sort((a, b) => {
      const as = regionStatsByStore.get(a.id) ?? emptyStats;
      const bs = regionStatsByStore.get(b.id) ?? emptyStats;
      if (storeSort === "거리 순" && userLocation) {
        const ad = hasStoreCoordinates(a) ? distanceKm(userLocation, { latitude: a.latitude!, longitude: a.longitude! }) : Number.POSITIVE_INFINITY;
        const bd = hasStoreCoordinates(b) ? distanceKm(userLocation, { latitude: b.latitude!, longitude: b.longitude! }) : Number.POSITIVE_INFINITY;
        return ad - bd || a.storeName.localeCompare(b.storeName, "ko");
      }
      if (storeSort === "품목 많은 순") return bs.total - as.total;
      if (storeSort === "미완료 많은 순") return (bs.notStarted + bs.inProgress) - (as.notStarted + as.inProgress);
      return a.storeName.localeCompare(b.storeName, "ko") || `${a.storeAddress}`.localeCompare(`${b.storeAddress}`, "ko");
    });
  }, [regionStores, regionStatsByStore, storeSort, userLocation]);
  const visibleRegionStores = useMemo(() => sortedRegionStores.filter((store) => {
    const ownItems = regionItemsByStore.get(store.id) ?? [];
    const searchText = [
      store.storeName,
      store.storeAddress,
      ...ownItems.flatMap((item) => [item.itemNo, item.productName, item.barcode, item.companyName, item.companyManager, item.companyTel, item.martTel]),
    ].join(" ");
    if (!searchText.includes(storeQuery)) return false;
    const ownStats = regionStatsByStore.get(store.id) ?? emptyStats;
    if (filter === "미완료" && ownStats.completed >= ownStats.total) return false;
    if (filter === "미진열" && !ownItems.some((item) => item.abnormalStatus === "미진열")) return false;
    if (filter !== "전체" && filter !== "미완료" && filter !== "사진누락" && filter !== "미진열" && !ownItems.some((item) => item.status === filter)) return false;
    if (filter === "사진누락" && !photosReady) return false;
    if (filter === "사진누락" && ownStats.photoMissing === 0) return false;
    return true;
  }), [sortedRegionStores, storeQuery, regionItemsByStore, regionStatsByStore, filter, photosReady]);
  const assignedVisibleRegionStores = useMemo(
    () => visibleRegionStores.filter((store) => store.mapIncluded === true),
    [visibleRegionStores],
  );
  const assignedRegionStores = useMemo(
    () => sortedRegionStores.filter((store) => store.mapIncluded === true),
    [sortedRegionStores],
  );
  const canUseStoreMap = true;
  const assignmentVisibleStores = useMemo(() => {
    const query = storeQuery.trim();
    if (!query) return sortedRegionStores;
    return sortedRegionStores.filter((store) => {
      const ownItems = regionItemsByStore.get(store.id) ?? [];
      return [
        store.storeName,
        store.storeAddress,
        ...ownItems.flatMap((item) => [item.itemNo, item.productName, item.barcode]),
      ].join(" ").includes(query);
    });
  }, [sortedRegionStores, storeQuery, regionItemsByStore]);
  const selectedStore = stores.find((store) => store.id === selectedStoreId);
  const storeItems = useMemo(() => items.filter((item) => item.storeId === selectedStoreId), [items, selectedStoreId]);
  const selectedItem = items.find((item) => item.id === selectedItemId);
  const reusableFrontPhotos = useMemo(() => photos
    .filter((photo) => photo.type === "STORE_FRONT" && photo.id !== selectedStore?.frontPhotoId)
    .sort((a, b) => `${b.takenAt}`.localeCompare(`${a.takenAt}`)), [photos, selectedStore?.frontPhotoId]);
  const visibleStoreItems = useMemo(() => storeItems
    .filter((item) => `${item.itemNo} ${item.productName} ${item.barcode} ${item.companyManager} ${item.companyName} ${item.companyTel} ${item.martTel}`.includes(itemQuery))
    .filter((item) => {
      if (filter === "전체") return true;
      if (filter === "미완료") return item.status !== "완료";
      if (filter === "사진누락") return item.status === "완료" && requiredPhotoLabels(item, photos.filter((photo) => photo.storeId === item.storeId)).length > 0;
      if (filter === "미진열") return item.abnormalStatus === "미진열";
      return item.status === filter;
    }), [storeItems, itemQuery, filter, photos]);
  const barcodeModalItem = items.find((item) => item.id === barcodeModalItemId);
  const barcodeModalItems = visibleStoreItems.length ? visibleStoreItems : storeItems;
  useEffect(() => {
    setStoreStatusDraft(selectedStore?.frontPhotoId ? selectedStore.operatingStatus ?? "" : "");
    setStoreStatusMessage("");
  }, [selectedStore?.id, selectedStore?.frontPhotoId, selectedStore?.operatingStatus]);
  useEffect(() => {
    if (view === "workspace" && workspaceMode === "map" && !canUseStoreMap) setWorkspaceMode("list");
  }, [view, workspaceMode, canUseStoreMap]);
  useEffect(() => {
    let cancelled = false;
    getBarcodeIndex()
      .then((data) => {
        if (!cancelled) setBarcodeIndex(data as BarcodeImageIndex);
      })
      .catch(() => {
        if (!cancelled) setBarcodeIndex({});
      });
    return () => {
      cancelled = true;
    };
  }, []);
  useEffect(() => {
    if (initialLocationRequested.current) return;
    initialLocationRequested.current = true;
    void locateUser();
  }, []);

  useEffect(() => {
    viewRef.current = view;
  }, [view]);

  useEffect(() => {
    if (isBooting) return;
    const state = makeHistoryState();
    if (!historyReadyRef.current) {
      window.history.replaceState(state, "", window.location.href);
      window.history.pushState(state, "", window.location.href);
      historyReadyRef.current = true;
      lastHistoryViewRef.current = view;
      return;
    }
    if (restoringHistoryRef.current) {
      window.history.replaceState(state, "", window.location.href);
      lastHistoryViewRef.current = view;
      restoringHistoryRef.current = false;
      return;
    }
    if (lastHistoryViewRef.current !== view) {
      window.history.pushState(state, "", window.location.href);
      lastHistoryViewRef.current = view;
    } else {
      window.history.replaceState(state, "", window.location.href);
    }
  }, [isBooting, view, currentRegion, selectedStoreId, selectedItemId, workspaceMode, barcodeModalItemId]);
  const stats = useMemo(() => summarize(regionItems, photos), [regionItems, photos]);

  const makeHistoryState = (): AppHistoryState => ({
    app: "research-supporter",
    view,
    currentRegion,
    selectedStoreId,
    selectedItemId,
    workspaceMode,
    barcodeModalItemId,
  });

  const applyHistoryState = (state: AppHistoryState) => {
    restoringHistoryRef.current = true;
    setMenuOpen(false);
    setSummaryOpen(false);
    setContactStoreId("");
    setStorageOpen(false);
    setFrontPhotoPickerOpen(false);
    setBarcodeModalItemId(state.barcodeModalItemId ?? "");
    setSelectedStoreId(state.selectedStoreId ?? "");
    setSelectedItemId(state.selectedItemId ?? "");
    setWorkspaceMode(state.workspaceMode ?? "list");
    if (state.currentRegion !== undefined) setSettingsState((old) => ({ ...old, currentRegion: state.currentRegion }));
    setView(state.view);
  };

  const showExitToast = () => {
    setAppBackToast("한 번 더 누르면 종료됩니다.");
    if (exitToastTimerRef.current) window.clearTimeout(exitToastTimerRef.current);
    exitToastTimerRef.current = window.setTimeout(() => setAppBackToast(""), 1800);
  };

  function askConfirm(options: ConfirmState) {
    return new Promise<boolean>((resolve) => {
      confirmResolver.current = resolve;
      setConfirmState(options);
    });
  }

  function closeConfirm(value: boolean) {
    confirmResolver.current?.(value);
    confirmResolver.current = null;
    setConfirmState(null);
  }

  async function refresh(region = currentRegion) {
    const [nextSettings, nextRegions, allStores, allItems] = await Promise.all([getSettings(), getRegions(), getStores(), getItems()]);
    const photoRegion = region ?? nextSettings.currentRegion;
    const nextPhotos = photoRegion ? await getPhotosByRegion(photoRegion) : [];
    setSettingsState(nextSettings);
    setRegions(nextRegions);
    setStores(allStores);
    setItems(allItems);
    setPhotos(nextPhotos);
    setPhotosReady(true);
    if (nextRegions.length && view === "upload") setView("regions");
    return { settings: nextSettings, regions: nextRegions, stores: allStores, items: allItems };
  }

  function locateUser(options: { force?: boolean; focus?: boolean } = {}) {
    if (locatePromiseRef.current && !options.force) return locatePromiseRef.current;
    const request = new Promise<{ latitude: number; longitude: number } | null>((resolve) => {
      if (!navigator.geolocation) {
        locatePromiseRef.current = null;
        resolve(null);
        return;
      }
      navigator.geolocation.getCurrentPosition(
        (position) => {
          const next = { latitude: position.coords.latitude, longitude: position.coords.longitude };
          setUserLocation(next);
          if (options.focus) setLocationFocusTick((value) => value + 1);
          locatePromiseRef.current = null;
          resolve(next);
        },
        () => {
          locatePromiseRef.current = null;
          resolve(null);
        },
        { enableHighAccuracy: options.force ?? false, timeout: options.force ? 6000 : 3000, maximumAge: options.force ? 0 : 300000 },
      );
    });
    locatePromiseRef.current = request;
    return request;
  }

  async function geocodeStores(targetStores: SurveyStore[], modeLabel: string) {
    const targets = targetStores.filter((store) => store.storeAddress);
    if (!targets.length) {
      setGeocodeMessage("좌표를 검색할 매장이 없습니다.");
      return;
    }
    const estimatedSeconds = Math.max(1, targets.length);
    const estimatedText = estimatedSeconds >= 60 ? `약 ${Math.ceil(estimatedSeconds / 60)}분` : `약 ${estimatedSeconds}초`;
    const ok = await askConfirm({
      title: "매장 위치 검색",
      message: `${modeLabel} ${targets.length.toLocaleString()}개의 위치를 주소로 검색합니다.\n무료 주소검색 정책을 지키기 위해 1초에 1개씩 처리하므로 ${estimatedText} 정도 걸릴 수 있습니다.\n계속할까요?`,
      confirmText: "시작",
      cancelText: "취소",
      plain: true,
    });
    if (!ok) return;
    setGeocoding(true);
    let success = 0;
    let failed = 0;
    try {
      for (const [index, store] of targets.entries()) {
        setGeocodeMessage(`좌표 검색 중 ${index + 1}/${targets.length}: ${mapSearchAddress(store.storeAddress)}`);
        try {
          const result = await geocodeAddress(store.storeAddress);
          const nextStore = result
            ? { ...store, latitude: result.latitude, longitude: result.longitude, geocodeStatus: "성공" as const, geocodedAt: now(), updatedAt: now() }
            : { ...store, geocodeStatus: "실패" as const, geocodedAt: now(), updatedAt: now() };
          await putStore(nextStore);
          setStores((old) => old.map((candidate) => candidate.id === store.id ? nextStore : candidate));
          if (result) success += 1;
          else failed += 1;
        } catch (error) {
          console.error(error);
          failed += 1;
          const failedStore = { ...store, geocodeStatus: "실패" as const, geocodedAt: now(), updatedAt: now() };
          await putStore(failedStore);
          setStores((old) => old.map((candidate) => candidate.id === store.id ? failedStore : candidate));
        }
        if (index < targets.length - 1) await delay(1100);
      }
      setGeocodeMessage(`좌표 검색 완료: 성공 ${success}개 · 실패 ${failed}개`);
    } finally {
      setGeocoding(false);
    }
  }

  useEffect(() => {
    refresh()
      .finally(() => setIsBooting(false));
  }, []);

  useEffect(() => {
    const warm = () => {
      void getPriceOcrWorker().catch(() => undefined);
    };
    const idleWindow = window as Window & {
      requestIdleCallback?: (callback: () => void, options?: { timeout?: number }) => number;
      cancelIdleCallback?: (id: number) => void;
    };
    if (idleWindow.requestIdleCallback) {
      const idleId = idleWindow.requestIdleCallback(warm, { timeout: 4000 });
      return () => idleWindow.cancelIdleCallback?.(idleId);
    }
    const timeoutId = globalThis.setTimeout(warm, 1200);
    return () => globalThis.clearTimeout(timeoutId);
  }, []);

  useEffect(() => {
    const updateVisualWidth = () => {
      const visualWidth = window.visualViewport?.width || Number.POSITIVE_INFINITY;
      const layoutWidth = document.documentElement.clientWidth || window.innerWidth || Number.POSITIVE_INFINITY;
      const width = Math.floor(Math.min(visualWidth, layoutWidth));
      if (Number.isFinite(width) && width > 0) {
        document.documentElement.style.setProperty("--app-visual-width", `${width}px`);
      }
    };
    updateVisualWidth();
    window.addEventListener("resize", updateVisualWidth);
    window.visualViewport?.addEventListener("resize", updateVisualWidth);
    window.visualViewport?.addEventListener("scroll", updateVisualWidth);
    return () => {
      window.removeEventListener("resize", updateVisualWidth);
      window.visualViewport?.removeEventListener("resize", updateVisualWidth);
      window.visualViewport?.removeEventListener("scroll", updateVisualWidth);
    };
  }, [view]);

  useEffect(() => {
    window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }, [view, currentRegion, selectedStoreId, selectedItemId]);

  useEffect(() => {
    const lockHorizontalScroll = () => {
      if (window.scrollX !== 0) window.scrollTo(0, window.scrollY);
    };
    window.addEventListener("scroll", lockHorizontalScroll, { passive: true });
    window.visualViewport?.addEventListener("scroll", lockHorizontalScroll);
    return () => {
      window.removeEventListener("scroll", lockHorizontalScroll);
      window.visualViewport?.removeEventListener("scroll", lockHorizontalScroll);
    };
  }, []);

  useEffect(() => {
    if (storeSort === "거리 순" && !userLocation) {
      setStoreSort("이름 순");
    }
  }, [storeSort, userLocation]);

  useEffect(() => {
    if (!menuOpen) return;
    const closeOnOutside = (event: PointerEvent) => {
      if (topbarRef.current?.contains(event.target as Node)) return;
      setMenuOpen(false);
    };
    document.addEventListener("pointerdown", closeOnOutside);
    return () => document.removeEventListener("pointerdown", closeOnOutside);
  }, [menuOpen]);

  useEffect(() => {
    if (view !== "workspace" || !selectedStoreId) return;
    window.setTimeout(() => document.getElementById(`store-card-${selectedStoreId}`)?.scrollIntoView({ block: "center", behavior: "smooth" }), 80);
  }, [view, selectedStoreId, visibleRegionStores.length]);

  async function updateSettings(patch: Partial<AppSettings>) {
    const next = { ...settings, ...patch };
    setSettingsState(next);
    await saveSettings(next);
  }

  async function analyzeUploadedFiles() {
    if (!surveyFile || !contactFile) {
      setUploadMessage("조사표와 업체 연락처 엑셀 파일을 선택해 주세요.");
      return;
    }
    if (regions.length || stores.length || items.length || photos.length) {
      const ok = await askConfirm({
        title: "자료 다시 구성",
        message: "새 엑셀 파일로 지역, 매장, 물품 목록을 다시 구성합니다. 기존 현장 입력값과 사진은 초기화될 수 있습니다. 계속할까요?",
        confirmText: "계속",
        cancelText: "취소",
        plain: true,
      });
      if (!ok) return;
    }
    setIsAnalyzing(true);
    setUploadMessage("자료를 분석하고 있습니다.");
    await new Promise((resolve) => setTimeout(resolve, 50));
    try {
      const parsed = await parseSurveyWorkbook(surveyFile);
      let parsedItems = parsed.items;
      let matched = 0;
      const before = parsedItems.filter((item) => item.companyTel).length;
      parsedItems = mergeContacts(parsedItems, await parseContactRows(contactFile));
      matched = parsedItems.filter((item) => item.companyTel).length - before;
      const rebuilt = rebuildStoresAndRegions(parsedItems);
      parsedItems = rebuilt.items;
      const parsedStores = rebuilt.stores.map((store) => {
        const first = parsedItems.find((item) => item.storeId === store.id);
        return first ? { ...store, storeAddress: first.storeAddress || store.storeAddress, storeName: first.storeName || store.storeName } : store;
      });
      let nextBarcodeIndex: BarcodeImageIndex = {};
      let extractedBarcodes = 0;
      if (barcodeFile) {
        setUploadMessage("바코드 이미지를 추출하고 있습니다.");
        nextBarcodeIndex = await extractBarcodeImages(barcodeFile);
        extractedBarcodes = Object.keys(nextBarcodeIndex).length;
      }
      await clearAllData();
      await saveParsedData(rebuilt.regions, parsedStores, parsedItems);
      await saveBarcodeIndex(nextBarcodeIndex);
      setBarcodeIndex(nextBarcodeIndex);
      setUploadMessage(`분석 완료: 전체 품목 ${parsedItems.length.toLocaleString()}개 / 지역 ${rebuilt.regions.length}개 / 매장 ${parsedStores.length}개 / 연락처 매칭 ${Math.max(0, matched)}개 / 바코드 이미지 ${extractedBarcodes.toLocaleString()}개`);
      await refresh(undefined);
      setView("regions");
    } catch (error) {
      console.error(error);
      setUploadMessage("자료를 분석하지 못했습니다. 엑셀 파일 양식을 확인해 주세요.");
    } finally {
      setIsAnalyzing(false);
    }
  }

  async function chooseRegion(region: string) {
    const nextSettings = { ...settings, currentRegion: region };
    setSettingsState(nextSettings);
    setSelectedStoreId("");
    setSelectedItemId("");
    setMapFocusStoreId("");
    setStoreQuery("");
    setItemQuery("");
    setFilter("전체");
    setStoreSort("이름 순");
    setWorkspaceMode("list");
    setPhotosReady(false);
    setPhotos([]);
    setView("workspace");
    void locateUser({ force: true });
    saveSettings(nextSettings).then(() => refresh(region));
  }

  async function openAssignment(region: string) {
    const nextSettings = { ...settings, currentRegion: region };
    setSettingsState(nextSettings);
    setSelectedStoreId("");
    setSelectedItemId("");
    setMapFocusStoreId("");
    setStoreQuery("");
    setItemQuery("");
    setWorkspaceMode("list");
    setStoreSort("이름 순");
    setWorkspaceToolsOpen(false);
    setPhotosReady(false);
    setPhotos([]);
    setView("assignment");
    saveSettings(nextSettings).then(() => refresh(region));
  }

  async function finishAssignment() {
    const assigned = regionStores.filter((store) => store.mapIncluded === true);
    const missingCoordinates = assigned.filter((store) => !hasStoreCoordinates(store)).length;
    if (missingCoordinates > 0) {
      const ok = await askConfirm({
        title: "위치정보 확인",
        message: `담당매장 중 위치정보가 없는 매장이 ${missingCoordinates.toLocaleString()}개 있습니다.\n위치정보를 가져오지 않으면 매장지도와 거리순 정렬을 사용할 수 없거나 일부 매장이 보이지 않을 수 있습니다.\n그래도 메인으로 돌아갈까요?`,
        confirmText: "돌아가기",
        cancelText: "계속 설정",
        plain: true,
      });
      if (!ok) return;
    }
    setStoreQuery("");
    setView("regions");
  }

  async function openStore(store: SurveyStore) {
    const todayValue = today();
    if (store.surveyDate !== todayValue) {
      const nextStore = { ...store, surveyDate: todayValue, updatedAt: now() };
      setStores((old) => old.map((candidate) => candidate.id === store.id ? nextStore : candidate));
      putStore(nextStore).then(() => refresh(store.region));
    }
    setSelectedStoreId(store.id);
    setItemQuery("");
    const nextSettings = { ...settings, lastOpenedStoreId: store.id, currentRegion: store.region };
    setSettingsState(nextSettings);
    setView("store");
    saveSettings(nextSettings);
  }

  function openStoreMap() {
    setMapFocusStoreId("");
    setWorkspaceMode("map");
  }

  function openStoreOnMap(store: SurveyStore) {
    setMapFocusStoreId(store.id);
    setWorkspaceMode("map");
  }

  async function setStoreAssigned(store: SurveyStore, assigned: boolean) {
    const nextStore = { ...store, mapIncluded: assigned, updatedAt: now() };
    await putStore(nextStore);
    setStores((old) => old.map((candidate) => candidate.id === store.id ? nextStore : candidate));
  }

  async function setStoresAssigned(targetStores: SurveyStore[], assigned: boolean) {
    const updated = targetStores.map((store) => ({ ...store, mapIncluded: assigned, updatedAt: now() }));
    await Promise.all(updated.map(putStore));
    const updates = new Map(updated.map((store) => [store.id, store]));
    setStores((old) => old.map((store) => updates.get(store.id) ?? store));
  }

  async function saveStorePhoto(file: File) {
    if (!selectedStore) return;
    if (selectedStore.frontPhotoId) await deletePhoto(selectedStore.frontPhotoId);
    const resized = await resizePhoto(file);
    const photo: SurveyPhoto = { id: uid("photo"), region: selectedStore.region, storeId: selectedStore.id, type: "STORE_FRONT", blob: resized.blob, originalName: file.name, mimeType: resized.mimeType, takenAt: now() };
    const nextStore = { ...selectedStore, frontPhotoId: photo.id, operatingStatus: selectedStore.frontPhotoId ? selectedStore.operatingStatus : undefined, status: "진행중" as const, startedAt: selectedStore.startedAt ?? now(), updatedAt: now() };
    await putPhoto(photo);
    await putStore(nextStore);
    await refresh(selectedStore.region);
  }

  async function useExistingStorePhoto(source: SurveyPhoto) {
    if (!selectedStore) return;
    if (source.id === selectedStore.frontPhotoId) {
      setFrontPhotoPickerOpen(false);
      return;
    }
    if (selectedStore.frontPhotoId) await deletePhoto(selectedStore.frontPhotoId);
    const copiedBlob = source.blob.slice(0, source.blob.size, source.mimeType || source.blob.type || "image/jpeg");
    const photo: SurveyPhoto = {
      id: uid("photo"),
      region: selectedStore.region,
      storeId: selectedStore.id,
      type: "STORE_FRONT",
      blob: copiedBlob,
      originalName: source.originalName ? `copy_${source.originalName}` : "store-front-copy.jpg",
      mimeType: source.mimeType || source.blob.type || "image/jpeg",
      takenAt: now(),
    };
    const nextStore = {
      ...selectedStore,
      frontPhotoId: photo.id,
      operatingStatus: selectedStore.frontPhotoId ? selectedStore.operatingStatus : undefined,
      status: "진행중" as const,
      startedAt: selectedStore.startedAt ?? now(),
      updatedAt: now(),
    };
    await putPhoto(photo);
    await putStore(nextStore);
    setFrontPhotoPickerOpen(false);
    await refresh(selectedStore.region);
  }

  async function removeStorePhoto() {
    if (!selectedStore?.frontPhotoId) return;
    await deletePhoto(selectedStore.frontPhotoId);
    await putStore({ ...selectedStore, frontPhotoId: undefined, operatingStatus: undefined, updatedAt: now() });
    await refresh(selectedStore.region);
  }

  async function saveItem(next: SurveyItem, photoOverride?: SurveyPhoto[]) {
    const storePhotos = photoOverride ?? await getPhotosByStore(next.storeId);
    const missing = requiredPhotoLabels(next, storePhotos);
    if (missing.length) {
      const label = next.normalDisplay === "X" ? "정상진열 X 품목" : next.normalDisplay === "O" ? "정상진열 품목" : "정상진열 여부가 선택되지 않아 기본 사진 기준";
      const ok = await askConfirm({
        title: "사진이 부족합니다",
        message: `${label}은 아래 사진이 필요합니다.\n\n- ${missing.join("\n- ")}\n\n그래도 저장하시겠습니까?`,
        confirmText: "사진 없이 저장",
        cancelText: "취소",
        danger: true,
      });
      if (!ok) return false;
    }
    const storeSurveyDate = stores.find((store) => store.id === next.storeId)?.surveyDate || today();
    const saved: SurveyItem = { ...next, surveyDate: storeSurveyDate, status: "완료", updatedAt: now() };
    await putItem(saved);
    await refresh(saved.region);
    return true;
  }

  function moveBarcodeModal(direction: -1 | 1) {
    const index = barcodeModalItems.findIndex((item) => item.id === barcodeModalItemId);
    const next = barcodeModalItems[index + direction];
    if (next) setBarcodeModalItemId(next.id);
  }

  function openItemFromBarcodeModal(itemId: string) {
    setBarcodeReturnItemId(itemId);
    setSelectedItemId(itemId);
    setBarcodeModalItemId("");
    setView("item");
  }

  function returnToBarcodeModal(itemId: string) {
    setSelectedItemId(itemId);
    setView("items");
    window.setTimeout(() => {
      document.getElementById(`item-card-${itemId}`)?.scrollIntoView({ block: "center", behavior: "smooth" });
      setBarcodeModalItemId(itemId);
    }, 80);
  }

  async function setStoreOperatingStatus(status: StoreOperatingStatus | "") {
    if (!selectedStore) return;
    if (!selectedStore.frontPhotoId) {
      await askConfirm({
        title: "매장 전경사진이 필요합니다",
        message: "매장 상태는 전경사진을 먼저 등록한 뒤 전환할 수 있습니다.",
        confirmText: "확인",
        cancelText: "닫기",
      });
      setStoreStatusDraft("");
      return;
    }
    if (!status) {
      await putStore({ ...selectedStore, operatingStatus: undefined, updatedAt: now() });
      await refresh(selectedStore.region);
      setStoreStatusMessage("상태가 미확인으로 변경되었습니다.");
      return;
    }
    if (status === selectedStore.operatingStatus) return;
    if (status === "영업 중" && (selectedStore.operatingStatus === "폐업" || selectedStore.operatingStatus === "임시휴업")) {
      await resetStoreItemsForOpen(selectedStore);
      return;
    }
    if (status === "폐업" || status === "임시휴업") {
      await applyStoreOperatingStatus(status);
      return;
    }
    await putStore({ ...selectedStore, operatingStatus: status, updatedAt: now() });
    await refresh(selectedStore.region);
    setStoreStatusMessage(`상태가 ${status}(으)로 변경되었습니다.`);
  }

  async function applyStoreOperatingStatus(status: Exclude<StoreOperatingStatus, "영업 중">) {
    if (!selectedStore) return;
    await applyStoreStatusToItems(selectedStore, status);
  }

  async function applyStoreStatusToItems(store: SurveyStore, status: Exclude<StoreOperatingStatus, "영업 중">) {
    const memoText = status === "폐업" ? "판매처 폐점" : "임시휴업";
    const ok = await askConfirm({
      title: `${status} 처리할까요?`,
      message: `${store.storeName} 하위 모든 품목을 일괄 변경합니다.\n\n이미 입력한 가격정보가 있어도 일괄로 데이터가 바뀝니다.`,
      confirmText: `${status} 처리`,
      cancelText: "취소",
      danger: true,
    });
    if (!ok) return;
    const storePhotos = store.id === selectedStore?.id ? photos : await getPhotosByStore(store.id);
    const removablePhotos = storePhotos.filter((photo) => photo.storeId === store.id && photo.type !== "STORE_FRONT");
    await Promise.all(removablePhotos.map((photo) => deletePhoto(photo.id)));
    const ownItems = store.id === selectedStore?.id ? storeItems : items.filter((item) => item.storeId === store.id);
    const surveyDate = store.surveyDate || today();
    const changedItems = ownItems.map((item) => ({
      ...item,
      surveyDate,
      normalDisplay: "X" as const,
      specMatch: "" as const,
      barcodeMatch: "" as const,
      normalPrice: null,
      hasDiscount: null,
      discountPrice: null,
      discountStartDate: "",
      discountEndDate: "",
      discountType: "",
      discountOral: false,
      discountPeriodMode: "" as const,
      barcodeRegistered: "X" as const,
      abnormalStatus: "미판매" as const,
      posChecked: "조회불가" as const,
      abnormalDisplay: "" as const,
      memo: appendMemoText(removeMemoTexts(item.memo, STORE_STATUS_MEMOS), memoText),
      status: "완료" as const,
      updatedAt: now(),
    }));
    await putStore({ ...store, operatingStatus: status, status: "완료", updatedAt: now() });
    await Promise.all(changedItems.map(putItem));
    await refresh(store.region);
    setStoreStatusMessage(`상태가 ${status}(으)로 변경되었습니다.`);
  }

  async function resetStoreItemsForOpen(store: SurveyStore) {
    const ok = await askConfirm({
      title: "영업 중으로 전환할까요?",
      message: `${store.storeName} 하위 품목의 가격정보와 품목 사진이 초기화됩니다.\n\n처음부터 다시 입력해야 합니다.`,
      confirmText: "전환",
      cancelText: "취소",
      danger: true,
    });
    if (!ok) return;
    const storePhotos = store.id === selectedStore?.id ? photos : await getPhotosByStore(store.id);
    const removablePhotos = storePhotos.filter((photo) => photo.storeId === store.id && photo.type !== "STORE_FRONT");
    await Promise.all(removablePhotos.map((photo) => deletePhoto(photo.id)));
    const ownItems = store.id === selectedStore?.id ? storeItems : items.filter((item) => item.storeId === store.id);
    const resetItems = ownItems.map((item) => ({
      ...item,
      surveyDate: "",
      normalDisplay: "" as const,
      specMatch: "" as const,
      barcodeMatch: "" as const,
      normalPrice: null,
      hasDiscount: null,
      discountPrice: null,
      discountStartDate: "",
      discountEndDate: "",
      discountType: "",
      discountOral: false,
      discountPeriodMode: "" as const,
      barcodeRegistered: "" as const,
      abnormalStatus: "" as const,
      posChecked: "" as const,
      posPrice: null,
      abnormalDisplay: "" as const,
      photoCase: "" as const,
      memo: "",
      status: "미조사" as const,
      updatedAt: now(),
    }));
    await putStore({ ...store, operatingStatus: "영업 중", status: "진행중", completedAt: undefined, updatedAt: now() });
    await Promise.all(resetItems.map(putItem));
    await refresh(store.region);
    setStoreStatusMessage("영업 중으로 전환하고 품목 정보를 초기화했습니다.");
  }

  async function doExportExcel(region = currentRegion) {
    if (!region) return;
    await exportRegionExcel(region, items.filter((item) => item.region === region));
  }

  async function doExportZip(region = currentRegion) {
    if (!region) return;
    const regionStoresForExport = stores.filter((store) => store.region === region);
    const regionItemsForExport = items.filter((item) => item.region === region);
    const regionPhotos = region === currentRegion ? photos : await getPhotosByRegion(region);
    await exportRegionZip(region, regionStoresForExport, regionItemsForExport, regionPhotos);
  }

  async function doBackup(region = currentRegion, all = false) {
    const scopeRegion = all ? undefined : region;
    const sourceStores = scopeRegion ? stores.filter((store) => store.region === scopeRegion) : stores;
    const sourceItems = scopeRegion ? items.filter((item) => item.region === scopeRegion) : items;
    const sourcePhotos = scopeRegion ? (scopeRegion === currentRegion ? photos : await getPhotosByRegion(scopeRegion)) : await getPhotos();
    const sourceRegions = scopeRegion ? regions.filter((candidate) => candidate.name === scopeRegion) : regions;
    await exportBackup(scopeRegion, sourceRegions, sourceStores, sourceItems, sourcePhotos, settings);
  }

  async function restoreBackup(file: File) {
    const payload = JSON.parse(await file.text()) as BackupPayload;
    const restoredPhotos = await Promise.all(payload.photos.map(async ({ dataUrl, ...photo }) => ({ ...photo, blob: await dataUrlToBlob(dataUrl) })));
    if (payload.scope === "all") {
      if (!confirm("현재 기기의 모든 자료와 입력값, 사진을 백업 파일 내용으로 덮어씁니다. 계속할까요?")) return;
      const nextSettings = { ...payload.settings, currentRegion: payload.settings.currentRegion ?? payload.regions[0]?.name };
      await importAllData(payload.regions, payload.stores, payload.items, restoredPhotos, nextSettings);
      await refresh(nextSettings.currentRegion);
      setView("regions");
      return;
    }
    const region = payload.region ?? payload.regions[0]?.name;
    if (!region) return;
    if (!confirm(`${region} 지역 데이터를 백업 파일 내용으로 덮어씁니다. 계속할까요?`)) return;
    await importRegionData(region, payload.stores, payload.items, restoredPhotos);
    await updateSettings({ currentRegion: region });
    await refresh(region);
    setView("regions");
  }

  async function openStorageInfo() {
    const estimate = await navigator.storage?.estimate?.();
    setStorageEstimate(estimate);
    setStorageOpen(true);
    setMenuOpen(false);
  }

  const regionSummary = (region: string, assignedOnly = false) => {
    const regionStoresForSummary = stores.filter((store) => store.region === region && (!assignedOnly || store.mapIncluded === true));
    const storeIds = new Set(regionStoresForSummary.map((store) => store.id));
    const regionItemsForSummary = items.filter((item) => item.region === region);
    const completed = regionStoresForSummary.filter((store) => {
      const own = regionItemsForSummary.filter((item) => item.storeId === store.id);
      return Boolean(store.frontPhotoId) && own.length > 0 && own.every((item) => item.status === "완료");
    }).length;
    const inProgress = regionStoresForSummary.filter((store) => {
      const own = regionItemsForSummary.filter((item) => item.storeId === store.id);
      const done = Boolean(store.frontPhotoId) && own.length > 0 && own.every((item) => item.status === "완료");
      return (Boolean(store.frontPhotoId) || own.some((item) => item.status === "완료" || item.status === "조사중")) && !done;
    }).length;
    return {
      total: regionStoresForSummary.length,
      completed,
      inProgress,
      notStarted: Math.max(0, regionStoresForSummary.length - completed - inProgress),
      photoMissing: summarize(regionItemsForSummary.filter((item) => storeIds.has(item.storeId)), region === currentRegion ? photos : []).photoMissing,
    };
  };
  const canGoBack = view !== "upload" && !(view === "regions" && regions.length > 0);
  const goBack = () => {
    setMenuOpen(false);
    if (view === "assignment") {
      setStoreQuery("");
      setView("regions");
    }
    else
    if (view === "workspace") {
      setStoreQuery("");
      setItemQuery("");
      setView("regions");
    }
    else if (view === "store") setView("workspace");
    else if (view === "items") {
      setItemQuery("");
      setView("store");
    }
    else if (view === "item") {
      if (barcodeReturnItemId && selectedItemId === barcodeReturnItemId) returnToBarcodeModal(selectedItemId);
      else setView("items");
    }
    else if (view === "validation") setView(currentRegion ? "workspace" : "regions");
    else if (view === "backup") setView(regions.length ? "regions" : "upload");
    else if (view === "regions") {
      setRegionQuery("");
      setStoreQuery("");
      setItemQuery("");
      setView("upload");
    }
  };
  useEffect(() => {
    const onPopState = (event: PopStateEvent) => {
      if (allowBrowserExitRef.current) return;
      const state = event.state as AppHistoryState | null;
      const rootView = viewRef.current === "regions" || viewRef.current === "upload";
      const targetRootView = state?.app === "research-supporter" && (state.view === "regions" || state.view === "upload");
      if (rootView && targetRootView) {
        const nowTime = Date.now();
        if (nowTime - lastExitBackRef.current < 1800) {
          allowBrowserExitRef.current = true;
          window.history.back();
          return;
        }
        lastExitBackRef.current = nowTime;
        showExitToast();
        window.history.pushState(makeHistoryState(), "", window.location.href);
        return;
      }
      if (state?.app === "research-supporter") {
        applyHistoryState(state);
        return;
      }

      if (rootView) {
        const nowTime = Date.now();
        if (nowTime - lastExitBackRef.current < 1800) {
          allowBrowserExitRef.current = true;
          window.history.back();
          return;
        }
        lastExitBackRef.current = nowTime;
        showExitToast();
        window.history.pushState(makeHistoryState(), "", window.location.href);
        return;
      }

      goBack();
      window.history.pushState(makeHistoryState(), "", window.location.href);
    };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, [view, currentRegion, selectedStoreId, selectedItemId, workspaceMode, barcodeModalItemId, regions.length, barcodeReturnItemId]);
  const screenTitle =
    view === "regions" ? "메인"
    : view === "assignment" ? "담당매장 관리"
    : view === "workspace" ? workspaceMode === "map" ? "매장지도" : "매장리스트"
    : view === "store" ? "매장정보"
    : view === "items" ? "물품리스트"
    : view === "item" ? "가격정보"
    : view === "validation" ? "검증"
    : view === "backup" ? "백업/복원"
    : "자료업로드";
  const menuAllRegionStats = useMemo(() => {
    if (!summaryOpen || view !== "regions") return emptyStats;
    const completed = regions.filter((region) => {
      const ownItems = items.filter((item) => item.region === region.name);
      return ownItems.length > 0 && ownItems.every((item) => item.status === "완료");
    }).length;
    return {
      total: regions.length,
      completed,
      inProgress: 0,
      notStarted: Math.max(0, regions.length - completed),
      photoMissing: 0,
    };
  }, [summaryOpen, view, regions, items]);
  if (isBooting) {
    return (
      <div className="app">
        <main className="boot-screen">
          <div className="loader-ring" aria-label="로딩 중" />
          <strong>가격조사 도우미</strong>
          <span>저장 데이터 확인 중</span>
        </main>
      </div>
    );
  }

  return (
    <div className="app">
      <header ref={topbarRef} className={`topbar ${menuOpen ? "menu-open" : ""}`}>
        <div className="top-main">
          <button className="top-back icon-button" onClick={goBack} disabled={!canGoBack} aria-label="뒤로가기">←</button>
          <div className="brand" aria-current="page">{screenTitle}</div>
          <span className="current" aria-hidden="true" />
          <button className="top-toggle icon-button" onClick={() => setMenuOpen((value) => !value)} aria-expanded={menuOpen} aria-label="메뉴 열기">
            <Menu size={20} />
          </button>
        </div>
        <div className="top-actions">
          <button onClick={() => { setStoreQuery(""); setItemQuery(""); setView("regions"); setMenuOpen(false); }}>HOME</button>
          <button onClick={() => { setSummaryOpen(true); setMenuOpen(false); }}>진행률 확인</button>
          <button disabled={!currentRegion} onClick={() => { setView("validation"); setMenuOpen(false); }}>검증</button>
          <button onClick={openStorageInfo}>자체저장공간</button>
          <button onClick={() => { setView("backup"); setMenuOpen(false); }}>백업/복원</button>
        </div>
      </header>
      {view === "upload" && (
        <main className="page narrow upload-page">
          <section className="upload-hero">
            <span>자료 업로드</span>
            <h1>조사 작업환경 만들기</h1>
            <p>조사표와 업체 연락처 엑셀 파일을 입력하면 지역, 매장, 물품 목록이 이 기기의 브라우저 저장공간에 생성됩니다. 서버 DB와 자동 동기화되지 않으므로 다른 기기에서 이어서 작업하려면 전체 백업 파일로 복원해 주세요.</p>
            <ul className="upload-notes">
              <li>조사표와 업체 연락처는 필수입니다.</li>
              <li>바코드 이미지 파일은 POS 확인용 바코드 모달에 사용됩니다.</li>
              <li>새 자료를 다시 분석하면 기존 입력 데이터가 초기화될 수 있으니 필요하면 먼저 전체 백업을 내려받아 주세요.</li>
            </ul>
          </section>
          <section className="panel upload-panel">
            <label className="file-card">
              <strong>1. 조사표 엑셀</strong>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => setSurveyFile(event.target.files?.[0] ?? null)} />
              <span>{surveyFile?.name ?? "선택된 파일 없음"}</span>
            </label>
            <label className="file-card">
              <strong>2. 업체 연락처 엑셀</strong>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => setContactFile(event.target.files?.[0] ?? null)} />
              <span>{contactFile?.name ?? "선택된 파일 없음"}</span>
            </label>
            <label className="file-card">
              <strong>3. 바코드 이미지 엑셀</strong>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => setBarcodeFile(event.target.files?.[0] ?? null)} />
              <span>{barcodeFile?.name ?? "선택하지 않으면 바코드 이미지는 표시되지 않습니다."}</span>
            </label>
            <button className="primary analyze-button" onClick={analyzeUploadedFiles} disabled={isAnalyzing}>{isAnalyzing ? "자료 분석 중..." : "자료 분석 시작"}</button>
            {uploadMessage && <p className="notice">{uploadMessage}</p>}
          </section>
          <button className="continue-button" disabled={!regions.length} onClick={() => setView("regions")}>지역리스트로 이동</button>
        </main>
      )}

      {view === "regions" && (
        <main className="page">
          <SearchBox value={regionQuery} onChange={setRegionQuery} placeholder="지역명 검색" />
          {currentRegion && regions.some((region) => region.name === currentRegion) && (
            <div className="recent-region">
              <div>
                <span>최근 지역</span>
                <button onClick={() => chooseRegion(currentRegion)}>{currentRegion}</button>
              </div>
              <a className="mini-map-link" target="_blank" href={TARGET_MAP_URL}>전체 지도</a>
            </div>
          )}
          <div className="grid">
            {regions.filter((region) => region.name.includes(regionQuery)).map((region) => {
              const summary = regionSummary(region.name);
              const assignedSummary = regionSummary(region.name, true);
              const regionStoreIds = new Set(stores.filter((store) => store.region === region.name).map((store) => store.id));
              const assignedStoreIds = new Set(stores.filter((store) => store.region === region.name && store.mapIncluded === true).map((store) => store.id));
              const regionPhotos = region.name === currentRegion ? photos : [];
              const allItemStats = summarize(items.filter((item) => item.region === region.name), regionPhotos);
              const assignedItemStats = summarize(items.filter((item) => item.region === region.name && assignedStoreIds.has(item.storeId)), regionPhotos);
              const hasPartialAssignment = regionStoreIds.size > 0 && assignedStoreIds.size !== regionStoreIds.size;
              return (
                <article className="card region-card" key={region.name}>
                  <div className="region-card-head">
                    <h2>{region.name}</h2>
                    <details className="card-menu subtle-menu">
                      <summary aria-label={`${region.name} 메뉴`}><MoreVertical size={18} /></summary>
                      <div className="menu-popover">
                        <button type="button" onClick={() => openAssignment(region.name)}>담당매장 관리</button>
                      </div>
                    </details>
                  </div>
                  <p className="area-summary">{region.areaSummary || region.city || "-"}</p>
                  <p className="muted">담당부서: {region.department || "-"}</p>
                  <RegionSummary
                    stats={summary.total ? summary : emptyStats}
                    itemStats={allItemStats}
                    assignedStats={hasPartialAssignment ? assignedSummary : undefined}
                    assignedItemStats={hasPartialAssignment ? assignedItemStats : undefined}
                  />
                  <div className="region-actions">
                    <button className="primary" onClick={() => chooseRegion(region.name)}>작업</button>
                    <button title="엑셀 내보내기" onClick={() => doExportExcel(region.name)}><Download size={16} />엑셀</button>
                    <button title="사진 ZIP" onClick={() => doExportZip(region.name)}><Download size={16} />사진</button>
                    <button title="백업 내려받기" onClick={() => doBackup(region.name)}><Download size={16} />백업</button>
                  </div>
                </article>
              );
            })}
          </div>
        </main>
      )}

      {view === "workspace" && currentRegion && (
        <main className="page workspace-page">
          <nav className="workspace-tabs" aria-label="매장 보기 방식">
            <button type="button" className={workspaceMode === "list" ? "active" : ""} onClick={() => setWorkspaceMode("list")}>매장 리스트</button>
            <button type="button" className={workspaceMode === "map" ? "active" : ""} onClick={() => canUseStoreMap && openStoreMap()} disabled={!canUseStoreMap}>매장 지도</button>
          </nav>
          {workspaceMode === "list" && (
            <div className="sticky-search workspace-search">
              <SearchBox value={storeQuery} onChange={setStoreQuery} placeholder="매장명 / 주소 / 품목명 / 품목코드 / 바코드" />
              <button className="tool-toggle" onClick={() => setWorkspaceToolsOpen((value) => !value)} aria-expanded={workspaceToolsOpen}>
                <SlidersHorizontal size={18} /> 필터 {workspaceToolsOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
              </button>
            </div>
          )}
          {workspaceToolsOpen && workspaceMode === "list" && (
            <section className="tool-panel">
              <div className="store-filter-sort-row">
                <FilterBar filter={filter} setFilter={setFilter} values={["전체", "미완료", "완료", "사진누락"]} />
                <label className="sort-control sort-only">
                <select value={storeSort} onChange={(event) => setStoreSort(event.target.value as StoreSort)}>
                  <option>이름 순</option>
                  <option>품목 많은 순</option>
                  <option>미완료 많은 순</option>
                  <option disabled={!userLocation}>거리 순</option>
                </select>
                </label>
              </div>
              {storeSort === "거리 순" && !assignedRegionStores.some(hasStoreCoordinates) && <p className="small-help warn">매장 위치정보가 없으면 거리순 정렬이 정확하지 않을 수 있습니다.</p>}
            </section>
          )}
          {workspaceMode === "list" && (
          <div className="list">
            {assignedVisibleRegionStores.map((store) => {
              const ownItems = regionItemsByStore.get(store.id) ?? [];
              const baseStats = regionStatsByStore.get(store.id) ?? emptyStats;
              const ownStats = photosReady ? baseStats : { ...baseStats, photoMissing: 0 };
              return (
                <StoreCard
                  key={store.id}
                  store={store}
                  stats={ownStats}
                  items={ownItems}
                  focused={selectedStoreId === store.id}
                  onOpen={() => openStore(store)}
                  onContacts={() => setContactStoreId(store.id)}
                  onAssignToggle={() => setStoreAssigned(store, store.mapIncluded !== true)}
                  onMapView={() => openStoreOnMap(store)}
                  distanceText={userLocation && hasStoreCoordinates(store) ? formatDistance(distanceKm(userLocation, { latitude: store.latitude!, longitude: store.longitude! })) : ""}
                />
              );
            })}
          </div>
          )}
          {workspaceMode === "map" && (
            <StoreMapView
              stores={assignedRegionStores}
              statsByStore={regionStatsByStore}
              userLocation={userLocation}
              locationFocusTick={locationFocusTick}
              focusStoreId={mapFocusStoreId}
              onOpen={(store) => openStore(store)}
              onContacts={(store) => setContactStoreId(store.id)}
              onToggle={(store) => setStoreAssigned(store, false)}
            />
          )}
          <button type="button" className="location-fab" onClick={() => { setMapFocusStoreId(""); locateUser({ force: true, focus: true }); }}>내 위치</button>
        </main>
      )}
      {view === "assignment" && currentRegion && (
        <main className="page">
          <div className="sticky-search workspace-search assignment-search">
            <SearchBox value={storeQuery} onChange={setStoreQuery} placeholder="매장명 / 주소 검색" />
          </div>
          <StoreAssignmentPanel
            stores={assignmentVisibleStores}
            totalStores={regionStores.length}
            statsByStore={regionStatsByStore}
            geocoding={geocoding}
            geocodeMessage={geocodeMessage}
            onGeocodeMissing={() => geocodeStores(regionStores.filter((store) => store.mapIncluded === true && !hasStoreCoordinates(store)), "위치정보가 없는 담당매장")}
            onGeocodeAll={() => geocodeStores(regionStores.filter((store) => store.mapIncluded === true), "담당매장")}
            onAssign={setStoreAssigned}
            onAssignAll={setStoresAssigned}
            onSave={finishAssignment}
          />
        </main>
      )}

      {view === "store" && selectedStore && (
        <main className="page narrow">
          <section className="panel">
            <h1>{selectedStore.storeName}</h1>
            <div className="store-address"><span>주소</span><strong>{selectedStore.storeAddress || "-"}</strong></div>
            <div className="store-address store-photo-heading"><span>매장 전경사진</span></div>
            {(() => {
              const frontPhoto = photos.find((photo) => photo.id === selectedStore.frontPhotoId);
              return (
            <div className={`photo-slot store-front-slot ${selectedStore.frontPhotoId ? "uploaded" : ""}`}>
              {frontPhoto && <PhotoPreview photo={frontPhoto} className="wide-preview" />}
              <div className="photo-actions store-front-actions">
                {!selectedStore.frontPhotoId && <PhotoInput label="촬영/첨부" onFile={saveStorePhoto} />}
                {selectedStore.frontPhotoId && <button className="danger" onClick={removeStorePhoto}>지우기</button>}
              </div>
              <button type="button" className="store-front-reuse" onClick={() => setFrontPhotoPickerOpen(true)}>기존 사진 사용</button>
            </div>
              );
            })()}
          </section>
          <section className="panel store-status-panel">
            <h2>상태</h2>
            <div className="store-operating">
              <span>현재 매장 상태</span>
              <strong className={`operating-badge ${selectedStore.frontPhotoId && selectedStore.operatingStatus ? operatingClass(selectedStore.operatingStatus) : "unknown"}`}>{storeDisplayStatus(selectedStore)}</strong>
            </div>
            <div className="store-state-actions">
              <select disabled={!selectedStore.frontPhotoId} value={selectedStore.frontPhotoId ? storeStatusDraft : ""} onChange={(event) => setStoreStatusDraft(event.target.value as StoreOperatingStatus | "")}>
                <option value="">미확인</option>
                <option value="영업 중">영업 중</option>
                <option value="폐업">폐업</option>
                <option value="임시휴업">임시휴업</option>
              </select>
              <button type="button" className="primary" disabled={!selectedStore.frontPhotoId} onClick={() => setStoreOperatingStatus(storeStatusDraft)}>저장</button>
            </div>
            {!selectedStore.frontPhotoId && <p className="small-help warn">매장 전경사진을 먼저 등록해야 상태를 전환할 수 있습니다.</p>}
            {selectedStore.frontPhotoId && !selectedStore.operatingStatus && <p className="small-help warn">조사 입력 전 매장 상태를 영업 중, 폐업, 임시휴업 중 하나로 설정해 주세요.</p>}
            {storeStatusMessage && <p className="ok store-status-message">{storeStatusMessage}</p>}
          </section>
          <Contacts items={storeItems} />
          <section className="panel">
            <p>조사 품목: {storeItems.length.toLocaleString()}건</p>
            <label className="store-date-row"><span>방문 조사일</span><input type="date" value={selectedStore.surveyDate} onChange={async (event) => { await putStore({ ...selectedStore, surveyDate: event.target.value, updatedAt: now() }); await refresh(selectedStore.region); }} /></label>
            <button className="primary sticky-lite" onClick={() => selectedStore.frontPhotoId && selectedStore.operatingStatus ? (setItemQuery(""), setView("items")) : alert(selectedStore.frontPhotoId ? "매장 상태를 먼저 설정해 주세요." : "매장 전경사진을 먼저 촬영/선택해 주세요.")}>조사 입력</button>
          </section>
        </main>
      )}

      {view === "items" && selectedStore && (
        <main className="page">
          <div className="sticky-search item-search">
            <SearchBox value={itemQuery} onChange={setItemQuery} placeholder="품목명 / 바코드 / 품목코드 / 담당자" />
            <button className="tool-toggle" onClick={() => setItemToolsOpen((value) => !value)} aria-expanded={itemToolsOpen}>
              <SlidersHorizontal size={18} /> 필터 {itemToolsOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
            </button>
          </div>
          {itemToolsOpen && (
            <section className="tool-panel">
              <Stats stats={summarize(storeItems, photos.filter((photo) => photo.storeId === selectedStore.id))} totalLabel="품목 전체" />
              <FilterBar filter={filter} setFilter={setFilter} values={["전체", "미완료", "완료", "사진누락", "미진열"]} />
            </section>
          )}
          <div className="list">
            {visibleStoreItems.map((item) => {
              const previewPhoto = photos.find((photo) => photo.itemId === item.id && ["PRODUCT_DISPLAY", "PRODUCT_INFO_BARCODE", "POS_RECEIPT"].includes(photo.type));
              const eligibility = getPriceEligibility(item);
              const itemPhotoMissing = item.status === "완료" && requiredPhotoLabels(item, photos.filter((photo) => photo.storeId === item.storeId)).length > 0;
              return (
                <article id={`item-card-${item.id}`} className={`card compact item-card ${selectedItemId === item.id ? "focused" : ""} ${item.status === "완료" ? "completed" : ""}`} key={item.id}>
                  <div className="item-card-head"><h2 className="item-title"><span className="item-code">{item.itemNo}</span><span>{item.productName}</span></h2><div className="item-badge-stack">{item.status !== "완료" && <Badge text={item.status} />}{itemPhotoMissing && <span className="badge badge-photo-missing">사진누락</span>}</div></div>
                  <div className={`item-card-body ${previewPhoto ? "" : "no-thumb"}`}>
                    <PhotoPreview photo={previewPhoto} className="item-thumb" />
                    <dl className="item-mini-info">
                      <dt>바코드</dt><dd>{item.barcode || "-"}</dd>
                      <dt>기준가격</dt><dd>{item.basePrice?.toLocaleString() ?? "-"}원</dd>
                      <dt>조사가격</dt><dd>{item.normalPrice?.toLocaleString() ?? "-"}원 {eligibility && <span className={`eligibility-badge ${eligibility.label === "부적격" ? "bad" : "good"}`} title={eligibility.reason}>{eligibility.label}</span>}</dd>
                    </dl>
                  </div>
                  <div className="item-card-actions">
                    <button type="button" onClick={() => setBarcodeModalItemId(item.id)}>바코드</button>
                    <button className="primary" onClick={() => { setBarcodeReturnItemId(""); setSelectedItemId(item.id); setView("item"); }}>입력</button>
                  </div>
                </article>
              );
            })}
          </div>
        </main>
      )}

      {view === "item" && selectedItem && (
        <ItemEditor item={selectedItem} storeItems={storeItems} storeOperatingStatus={stores.find((store) => store.id === selectedItem.storeId)?.operatingStatus ?? ""} photos={photos.filter((photo) => photo.storeId === selectedItem.storeId)} fromBarcodeFlow={barcodeReturnItemId === selectedItem.id} onSave={saveItem} onSaved={async () => { await refresh(selectedItem.region); }} onList={(focusId) => { if (focusId) setSelectedItemId(focusId); setView("items"); }} onStoreList={() => { setFilter("전체"); setSelectedStoreId(selectedItem.storeId); setView("workspace"); }} onMove={(id) => setSelectedItemId(id)} onBarcodeReturn={returnToBarcodeModal} askConfirm={askConfirm} />
      )}

      {view === "validation" && (
        <main className="page">
          <Validation title="미완료 품목" items={regionItems.filter((item) => item.status !== "완료")} open={(id) => { setSelectedItemId(id); setView("item"); }} />
          <Validation title="사진누락 품목" items={regionItems.filter((item) => requiredPhotoLabels(item, photos.filter((photo) => photo.storeId === item.storeId)).length > 0)} open={(id) => { setSelectedItemId(id); setView("item"); }} />
          <Validation title="정상진열 X 품목" items={regionItems.filter((item) => item.normalDisplay === "X")} open={(id) => { setSelectedItemId(id); setView("item"); }} />
          <Validation title="특이사항 입력 품목" items={regionItems.filter((item) => item.memo)} open={(id) => { setSelectedItemId(id); setView("item"); }} />
        </main>
      )}

      {view === "backup" && (
        <main className="page narrow">
          <section className="backup-grid">
            <article className="panel">
              <h2>백업 내려받기</h2>
              <p className="muted">현재 브라우저에 저장된 조사 데이터와 사진을 JSON으로 저장합니다.</p>
              <button className="primary full-button" onClick={() => doBackup(undefined, true)}><Download size={17} />전체 백업 다운로드</button>
            </article>
            <article className="panel">
              <h2>백업 업로드</h2>
              <p className="muted">다른 폰이나 PC에서 만든 백업 JSON을 불러옵니다.</p>
              <label className="photo-button"><Upload size={18} />백업 JSON 업로드<input type="file" accept="application/json,.json" onChange={(event) => event.target.files?.[0] && restoreBackup(event.target.files[0])} /></label>
            </article>
            <article className="panel">
              <h2>초기화</h2>
              <p className="muted">기기 안의 IndexedDB 조사 데이터를 모두 삭제합니다.</p>
            <button className="danger" onClick={async () => { if (confirm("모든 IndexedDB 데이터를 삭제합니다. 계속할까요?")) { await clearAllData(); await refresh(undefined); setView("upload"); } }}>전체 데이터 초기화</button>
            </article>
          </section>
        </main>
      )}
      {barcodeModalItem && (
        <BarcodePhotoModal
          item={barcodeModalItem}
          barcodeSrc={barcodeImageSrc(barcodeModalItem, barcodeIndex)}
          currentIndex={Math.max(0, barcodeModalItems.findIndex((item) => item.id === barcodeModalItem.id))}
          totalCount={barcodeModalItems.length}
          canPrev={barcodeModalItems.findIndex((item) => item.id === barcodeModalItem.id) > 0}
          canNext={barcodeModalItems.findIndex((item) => item.id === barcodeModalItem.id) < barcodeModalItems.length - 1}
          onOpenItem={() => openItemFromBarcodeModal(barcodeModalItem.id)}
          onPrev={() => moveBarcodeModal(-1)}
          onNext={() => moveBarcodeModal(1)}
          onClose={() => setBarcodeModalItemId("")}
        />
      )}
      {contactStoreId && (
        <ContactModal
          store={stores.find((store) => store.id === contactStoreId)}
          items={items.filter((item) => item.storeId === contactStoreId)}
          onClose={() => setContactStoreId("")}
        />
      )}
      {storageOpen && (
        <StorageModal
          estimate={storageEstimate}
          photoCount={photos.length}
          onRefresh={openStorageInfo}
          onClose={() => setStorageOpen(false)}
        />
      )}
      {frontPhotoPickerOpen && selectedStore && (
        <StoreFrontPhotoPicker
          photos={reusableFrontPhotos}
          stores={stores}
          onSelect={useExistingStorePhoto}
          onClose={() => setFrontPhotoPickerOpen(false)}
        />
      )}
      {confirmState && <ConfirmDialog state={confirmState} onClose={closeConfirm} />}
      {appBackToast && <div className="save-toast">{appBackToast}</div>}
      {summaryOpen && view === "regions" && (
        <SummaryModal
          region="전체 지역"
          stats={menuAllRegionStats}
          storeCount={regions.length}
          completedStoreCount={menuAllRegionStats.completed}
          mode="regions"
          onClose={() => setSummaryOpen(false)}
        />
      )}
      {summaryOpen && view === "workspace" && (
        <SummaryModal
          region={currentRegion}
          stats={stats}
          storeCount={regionStores.length}
          completedStoreCount={regionStores.filter((store) => {
            const ownStats = regionStatsByStore.get(store.id) ?? emptyStats;
            return Boolean(store.frontPhotoId) && ownStats.total > 0 && ownStats.completed === ownStats.total;
          }).length}
          mode="workspace"
          onClose={() => setSummaryOpen(false)}
        />
      )}
      {summaryOpen && view === "items" && selectedStore && (
        <SummaryModal
          region={selectedStore.storeName}
          stats={summarize(storeItems, photosByStore.get(selectedStore.id) ?? [])}
          storeCount={1}
          completedStoreCount={storeItems.every((item) => item.status === "완료") ? 1 : 0}
          mode="items"
          onClose={() => setSummaryOpen(false)}
        />
      )}
      {summaryOpen && view === "store" && selectedStore && (
        <SummaryModal
          region={selectedStore.storeName}
          stats={summarize(storeItems, photosByStore.get(selectedStore.id) ?? [])}
          storeCount={1}
          completedStoreCount={storeItems.length > 0 && storeItems.every((item) => item.status === "완료") ? 1 : 0}
          mode="items"
          onClose={() => setSummaryOpen(false)}
        />
      )}
      {summaryOpen && view === "item" && selectedStore && (
        <SummaryModal
          region={selectedStore.storeName}
          stats={summarize(storeItems, photosByStore.get(selectedStore.id) ?? [])}
          storeCount={1}
          completedStoreCount={storeItems.length > 0 && storeItems.every((item) => item.status === "완료") ? 1 : 0}
          mode="items"
          onClose={() => setSummaryOpen(false)}
        />
      )}
    </div>
  );
}

function SearchBox({ value, onChange, placeholder }: { value: string; onChange: (value: string) => void; placeholder: string }) {
  return <label className="search"><Search size={18} /><input value={value} onChange={(event) => onChange(event.target.value)} placeholder={placeholder} /></label>;
}

function PhotoPreview({ photo, className = "" }: { photo?: SurveyPhoto; className?: string }) {
  const [url, setUrl] = useState("");
  useEffect(() => {
    if (!photo) {
      setUrl("");
      return;
    }
    const nextUrl = URL.createObjectURL(photo.blob);
    setUrl(nextUrl);
    return () => URL.revokeObjectURL(nextUrl);
  }, [photo?.id]);
  if (!photo || !url) return null;
  return <img className={`photo-preview ${className}`} src={url} alt="업로드 사진 미리보기" loading="lazy" />;
}

function StoreFrontPhotoPicker({ photos, stores, onSelect, onClose }: { photos: SurveyPhoto[]; stores: SurveyStore[]; onSelect: (photo: SurveyPhoto) => void | Promise<void>; onClose: () => void }) {
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal front-photo-picker-modal">
        <div className="modal-head">
          <div>
            <h2>기존 전경사진 사용</h2>
            <p>이미 촬영한 전경사진을 선택하면 현재 매장 사진으로 복사됩니다.</p>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        {photos.length === 0 ? (
          <p className="small-help">사용할 수 있는 전경사진이 아직 없습니다.</p>
        ) : (
          <div className="front-photo-grid">
            {photos.map((photo) => {
              const store = stores.find((candidate) => candidate.id === photo.storeId);
              return (
                <button type="button" className="front-photo-option" key={photo.id} onClick={() => onSelect(photo)}>
                  <PhotoPreview photo={photo} className="front-photo-thumb" />
                  <strong>{store?.storeName || "매장 정보 없음"}</strong>
                  <span>{store?.storeAddress || "-"}</span>
                  <em>{photo.takenAt ? photo.takenAt.slice(0, 16).replace("T", " ") : "-"}</em>
                </button>
              );
            })}
          </div>
        )}
      </section>
    </div>
  );
}

function StorageModal({ estimate, photoCount, onRefresh, onClose }: { estimate?: StorageEstimate; photoCount: number; onRefresh: () => void; onClose: () => void }) {
  const used = estimate?.usage ?? 0;
  const quota = estimate?.quota ?? 0;
  const available = availableStorageBytes(estimate);
  const percent = quota ? Math.min(100, Math.round((used / quota) * 100)) : 0;
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal">
        <div className="modal-head">
          <div>
            <h2>자체저장공간</h2>
            <p>이 브라우저가 앱 데이터와 사진을 저장하는 공간입니다.</p>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        <div className="storage-meter">
          <div><strong>{photoCount.toLocaleString()}장</strong><span>저장 사진 수</span></div>
          <div><strong>{formatBytes(used)}</strong><span>사용 중</span></div>
          <div><strong>{formatBytes(available)}</strong><span>여유공간</span></div>
        </div>
        <div className="progress-line storage-progress"><span style={{ width: `${percent}%` }} /></div>
        <p className="small-help">브라우저가 알려준 저장공간 기준입니다. 실제 저장 가능 용량은 기기 여유공간과 브라우저 정책에 따라 달라질 수 있으니 조사 중에는 지역별 백업을 자주 내려받아 주세요.</p>
        <button onClick={onRefresh}>다시 확인</button>
      </section>
    </div>
  );
}

function BarcodePhotoModal({
  item,
  barcodeSrc,
  currentIndex,
  totalCount,
  canPrev,
  canNext,
  onOpenItem,
  onPrev,
  onNext,
  onClose,
}: {
  item: SurveyItem;
  barcodeSrc: string;
  currentIndex: number;
  totalCount: number;
  canPrev: boolean;
  canNext: boolean;
  onOpenItem: () => void;
  onPrev: () => void;
  onNext: () => void;
  onClose: () => void;
}) {
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal barcode-modal">
        <div className="modal-head">
          <div>
            <h2>바코드 / POS 사진</h2>
            <p>{totalCount ? `${currentIndex + 1}/${totalCount}` : "1/1"} · {item.itemNo} {item.productName}</p>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        <div className="barcode-modal-body">
          <div className="barcode-info-row">
            <span>바코드</span>
            <strong>{item.barcode || "-"}</strong>
          </div>
          {barcodeSrc ? (
            <div className="barcode-image-wrap">
              <img src={barcodeSrc} alt={`${item.productName} 바코드`} className="barcode-image" loading="lazy" />
            </div>
          ) : (
            <p className="small-help warn">등록된 바코드 이미지가 없습니다.</p>
          )}
          <p className="barcode-flow-note">POS에서 이 바코드를 확인한 뒤 입력을 누르면 ③ 상태 입력 위치로 이동합니다.</p>
          <button type="button" className="primary" onClick={onOpenItem}>입력</button>
          <div className="barcode-nav">
            <button type="button" disabled={!canPrev} onClick={onPrev}>이전</button>
            <button type="button" disabled={!canNext} onClick={onNext}>다음</button>
          </div>
        </div>
      </section>
    </div>
  );
}

function ConfirmDialog({ state, onClose }: { state: ConfirmState; onClose: (value: boolean) => void }) {
  const message = [state.title, state.message].filter(Boolean).join("\n");
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal confirm-dialog">
        <div className="confirm-message">
          {message.split("\n").map((line, index) => <p key={`${line}-${index}`}>{line || "\u00a0"}</p>)}
        </div>
        <div className="confirm-actions">
          <button onClick={() => onClose(false)}>{state.cancelText ?? "취소"}</button>
          <button className={state.danger ? "danger" : "primary"} onClick={() => onClose(true)}>{state.confirmText ?? "확인"}</button>
        </div>
      </section>
    </div>
  );
}

function FilterBar({ filter, setFilter, values = ["전체", "미완료", "미조사", "조사중", "완료", "사진누락"] }: { filter: Filter; setFilter: (filter: Filter) => void; values?: Filter[] }) {
  return <div className="segmented filter-chips">{values.map((value) => <button className={filter === value ? "active" : ""} key={value} onClick={() => setFilter(value)}>{value}</button>)}</div>;
}

function Stats({ stats, totalLabel = "전체" }: { stats: RegionStats; totalLabel?: string }) {
  return (
    <div className="stats">
      <div className="stats-line">
        <strong>{totalLabel} {stats.total.toLocaleString()}개</strong>
        <span>완료 {stats.completed.toLocaleString()}</span>
        <span>조사중 {stats.inProgress.toLocaleString()}</span>
        <span>미조사 {stats.notStarted.toLocaleString()}</span>
      </div>
      <div className="stats-progress"><span style={{ width: `${stats.total ? Math.round((stats.completed / stats.total) * 100) : 0}%` }} /></div>
      <div className={stats.photoMissing > 0 ? "stats-missing" : "stats-photo-ok"}>{stats.photoMissing > 0 ? `사진누락 ${stats.photoMissing.toLocaleString()}개` : "사진"}</div>
    </div>
  );
}

function RegionSummary({ stats, itemStats, assignedStats, assignedItemStats }: { stats: RegionStats; itemStats: RegionStats; assignedStats?: RegionStats; assignedItemStats?: RegionStats }) {
  const storePercent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;
  const itemPercent = itemStats.total ? Math.round((itemStats.completed / itemStats.total) * 100) : 0;
  return (
    <div className="region-summary">
      <div className="region-metric">
        <span>매장</span>
        <strong>{stats.completed.toLocaleString()}<small>/{stats.total.toLocaleString()}</small></strong>
        <div className="mini-progress"><i style={{ width: `${storePercent}%` }} /></div>
        {assignedStats ? <em className="assigned-progress">담당 {assignedStats.completed.toLocaleString()}/{assignedStats.total.toLocaleString()}</em> : <em>미조사 {stats.notStarted.toLocaleString()}</em>}
      </div>
      <div className="region-metric">
        <span>품목</span>
        <strong>{itemStats.completed.toLocaleString()}<small>/{itemStats.total.toLocaleString()}</small></strong>
        <div className="mini-progress"><i style={{ width: `${itemPercent}%` }} /></div>
        {assignedItemStats ? <em className="assigned-progress">담당 {assignedItemStats.completed.toLocaleString()}/{assignedItemStats.total.toLocaleString()}</em> : <em>미조사 {itemStats.notStarted.toLocaleString()}</em>}
      </div>
    </div>
  );
}

function Badge({ text }: { text: string }) {
  return <span className={`badge badge-${text}`}>{text}</span>;
}

function getSurveySalePrice(item: SurveyItem) {
  let price = item.normalPrice;
  const longDiscount = item.hasDiscount && item.discountType.replace("구두", "") === "②" && item.discountPrice !== null;
  if (longDiscount) price = item.discountPrice;
  if (item.memo.includes("1+1") && price !== null) price = Math.round(price / 2);
  return price;
}

function getPriceEligibility(item: SurveyItem) {
  if (item.abnormalStatus === "미판매") return { label: "부적격", reason: "미판매" };
  if (item.abnormalStatus === "미진열") return { label: "부적격", reason: "미진열" };
  if (item.barcodeRegistered === "X") return { label: "부적격", reason: "바코드 미등록" };
  if (item.memo.includes("판매처 폐점")) return { label: "부적격", reason: "폐업" };
  if (item.memo.includes("임시휴업")) return { label: "부적격", reason: "임시휴업" };
  const salePrice = getSurveySalePrice(item);
  if (item.basePrice === null || salePrice === null) return undefined;
  if (salePrice < item.basePrice) return { label: "부적격", reason: "저가판매" };
  return { label: "적격", reason: "정상 판매 및 정상 가격" };
}

function operatingClass(status?: StoreOperatingStatus) {
  return `operating-${(status ?? "영업 중").replace(/\s/g, "")}`;
}

function storeDisplayStatus(store: SurveyStore) {
  return store.frontPhotoId ? store.operatingStatus ?? "미확인" : "미확인";
}

function hasStoreCoordinates(store: SurveyStore) {
  return typeof store.latitude === "number" && Number.isFinite(store.latitude) && typeof store.longitude === "number" && Number.isFinite(store.longitude);
}

function isStoreComplete(store: SurveyStore, stats: RegionStats) {
  return Boolean(store.frontPhotoId) && stats.total > 0 && stats.completed === stats.total;
}

function StoreMoreMenu({ store, onAssignToggle, onMapView }: { store: SurveyStore; onAssignToggle: () => void; onMapView?: () => void }) {
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: PointerEvent) => {
      if (menuRef.current?.contains(event.target as Node)) return;
      setOpen(false);
    };
    document.addEventListener("pointerdown", close);
    return () => document.removeEventListener("pointerdown", close);
  }, [open]);
  const closeAfter = (action?: () => void) => {
    action?.();
    setOpen(false);
  };
  return (
    <div className={`card-menu ${open ? "open" : ""}`} ref={menuRef}>
      <button type="button" className="card-menu-trigger" aria-label="매장 메뉴" aria-expanded={open} onClick={() => setOpen((value) => !value)}><MoreVertical size={18} /></button>
      {open && (
        <div className="menu-popover">
          {onMapView && hasStoreCoordinates(store) && <button type="button" onClick={() => closeAfter(onMapView)}>매장지도에서 보기</button>}
          <button type="button" onClick={() => closeAfter(onAssignToggle)}>{store.mapIncluded === true ? "담당매장 제외" : "담당매장 포함"}</button>
          {mapLinks(store.storeAddress).map(([name, href]) => <a key={name} href={href} target="_blank" onClick={() => setOpen(false)}>{name} 지도 보기</a>)}
        </div>
      )}
    </div>
  );
}

function StoreCard({
  store,
  stats,
  items,
  focused,
  onOpen,
  onContacts,
  onAssignToggle,
  onMapView,
  distanceText,
}: {
  store: SurveyStore;
  stats: RegionStats;
  items: SurveyItem[];
  focused: boolean;
  onOpen: () => void;
  onContacts: () => void;
  onAssignToggle: () => void;
  onMapView: () => void;
  distanceText?: string;
}) {
  const completed = items.filter((item) => item.status === "완료");
  const latestSurveyDate = completed.map((item) => item.surveyDate).filter(Boolean).sort().at(-1) ?? "-";
  const percent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;
  const completedStore = isStoreComplete(store, stats);
  const displayOperatingStatus = storeDisplayStatus(store);

  return (
    <article id={`store-card-${store.id}`} className={`card store-card ${focused ? "focused" : ""} ${completedStore ? "completed" : ""}`}>
      <div className="card-head">
        <div>
          <h2 className="store-card-title" title={store.storeName}><span className="store-name-text">{store.storeName}</span><span className={`operating-badge small ${displayOperatingStatus === "미확인" ? "unknown" : operatingClass(displayOperatingStatus as StoreOperatingStatus)}`}>{displayOperatingStatus}</span></h2>
          <p>{store.storeAddress || "주소 없음"}</p>
        </div>
        <StoreMoreMenu store={store} onAssignToggle={onAssignToggle} onMapView={onMapView} />
      </div>
      {store.mapIncluded !== true && <span className="map-excluded-badge">담당 미선택</span>}
      <div className="store-progress">
        <div className="store-metric-row">
          <span>물품</span>
          <strong>{stats.completed.toLocaleString()}<small>/{stats.total.toLocaleString()}</small></strong>
        </div>
        <div className="progress-line"><span style={{ width: `${percent}%` }} /></div>
      </div>
      <div className="store-meta">
        <span className={`store-distance ${distanceText ? "" : "empty"}`}>{distanceText ? `현재 위치 ${distanceText}` : "현재 위치 -"}</span>
        {stats.photoMissing > 0 && <span className="store-missing">사진 누락 {stats.photoMissing.toLocaleString()}건</span>}
        <span className="store-date">조사일: {latestSurveyDate}</span>
      </div>
      <div className="card-actions">
        <button onClick={onContacts}>담당자 정보</button>
        <button className="primary" onClick={onOpen}>입력</button>
      </div>
    </article>
  );
}

function StoreMapView({ stores, statsByStore, userLocation, locationFocusTick, focusStoreId, onOpen, onContacts, onToggle }: { stores: SurveyStore[]; statsByStore: Map<string, RegionStats>; userLocation: { latitude: number; longitude: number } | null; locationFocusTick: number; focusStoreId: string; onOpen: (store: SurveyStore) => void; onContacts: (store: SurveyStore) => void; onToggle: (store: SurveyStore) => void | Promise<void> }) {
  const mapNode = useRef<HTMLDivElement | null>(null);
  const leafletMap = useRef<import("leaflet").Map | null>(null);
  const markerLayer = useRef<import("leaflet").LayerGroup | null>(null);
  const mappedStores = useMemo(() => stores.filter(hasStoreCoordinates), [stores]);
  const [activeStoreId, setActiveStoreId] = useState("");
  const [mapReady, setMapReady] = useState(0);
  const activeStore = stores.find((store) => store.id === activeStoreId);
  const completedCount = stores.filter((store) => isStoreComplete(store, statsByStore.get(store.id) ?? emptyStats)).length;
  const locationText = userLocation ? `내 위치: ${userLocation.latitude.toFixed(5)}, ${userLocation.longitude.toFixed(5)}` : "내 위치: 확인 필요";
  const mapSignature = [
    mappedStores.map((store) => `${store.id}:${store.latitude}:${store.longitude}:${isStoreComplete(store, statsByStore.get(store.id) ?? emptyStats) ? "1" : "0"}`).join("|"),
    userLocation ? `${userLocation.latitude}:${userLocation.longitude}` : "",
  ].join("::");

  useEffect(() => {
    let cancelled = false;
    let fallbackTimer: ReturnType<typeof window.setTimeout> | null = null;
    import("leaflet").then((leaflet) => {
      if (cancelled || !mapNode.current || leafletMap.current) return;
      const map = leaflet.map(mapNode.current, {
        zoomControl: true,
        attributionControl: true,
        preferCanvas: true,
      });
      leafletMap.current = map;
      const tileConfigs: Array<{ url: string; options: import("leaflet").TileLayerOptions }> = [
        {
          url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
          options: { maxZoom: 19, attribution: "&copy; OpenStreetMap contributors" },
        },
        {
          url: "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}.png",
          options: { maxZoom: 19, subdomains: ["a", "b", "c", "d"], attribution: "&copy; OpenStreetMap contributors &copy; CARTO" },
        },
        {
          url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}",
          options: { maxZoom: 13, attribution: "Tiles &copy; Esri, OpenStreetMap contributors" },
        },
      ];
      let tileIndex = 0;
      let activeTiles: import("leaflet").TileLayer | null = null;
      const addTiles = () => {
        if (cancelled) return;
        const config = tileConfigs[Math.min(tileIndex, tileConfigs.length - 1)];
        activeTiles?.remove();
        if (fallbackTimer) window.clearTimeout(fallbackTimer);
        activeTiles = leaflet.tileLayer(config.url, config.options).addTo(map);
        fallbackTimer = window.setTimeout(() => {
          if (tileIndex >= tileConfigs.length - 1) return;
          tileIndex += 1;
          addTiles();
        }, 2800);
        activeTiles.once("tileload", () => {
          if (fallbackTimer) window.clearTimeout(fallbackTimer);
          fallbackTimer = null;
        });
        activeTiles.on("tileerror", () => {
          if (tileIndex >= tileConfigs.length - 1) return;
          tileIndex += 1;
          addTiles();
        });
      };
      addTiles();
      map.setView(userLocation ? [userLocation.latitude, userLocation.longitude] : [37.5665, 126.978], userLocation ? 15 : 12);
      setMapReady((value) => value + 1);
      window.requestAnimationFrame(() => map.invalidateSize());
      window.setTimeout(() => map.invalidateSize(), 120);
      window.setTimeout(() => map.invalidateSize(), 500);
    });
    return () => {
      cancelled = true;
      if (fallbackTimer) window.clearTimeout(fallbackTimer);
      markerLayer.current?.remove();
      leafletMap.current?.remove();
      markerLayer.current = null;
      leafletMap.current = null;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    import("leaflet").then((leaflet) => {
      const map = leafletMap.current;
      if (cancelled || !map || !mapReady) return;
      markerLayer.current?.remove();
      const layer = leaflet.layerGroup().addTo(map);
      markerLayer.current = layer;
      const bounds: import("leaflet").LatLngExpression[] = [];
      mappedStores.forEach((store) => {
        const stats = statsByStore.get(store.id) ?? emptyStats;
        const completed = isStoreComplete(store, stats);
        const latLng: import("leaflet").LatLngExpression = [store.latitude!, store.longitude!];
        bounds.push(latLng);
        leaflet.circleMarker(latLng, {
          radius: 9,
          color: completed ? "#6b7280" : "#dc2626",
          weight: 2,
          fillColor: completed ? "#9ca3af" : "#ef4444",
          fillOpacity: 0.82,
        })
          .addTo(layer)
          .bindTooltip(`${store.storeName} · ${completed ? "완료" : "미완료"}`)
          .on("click", () => {
            setActiveStoreId(store.id);
          });
      });
      if (userLocation) {
        const latLng: import("leaflet").LatLngExpression = [userLocation.latitude, userLocation.longitude];
        bounds.push(latLng);
        leaflet.circleMarker(latLng, {
          radius: 8,
          color: "#0284c7",
          weight: 3,
          fillColor: "#38bdf8",
          fillOpacity: 0.9,
        }).addTo(layer).bindTooltip("내 위치");
      }
      window.requestAnimationFrame(() => map.invalidateSize());
    });
    return () => {
      cancelled = true;
    };
  }, [mapReady, mapSignature, mappedStores, statsByStore, userLocation]);

  useEffect(() => {
    if (activeStoreId && !mappedStores.some((store) => store.id === activeStoreId)) setActiveStoreId("");
  }, [mappedStores, activeStoreId]);

  useEffect(() => {
    const map = leafletMap.current;
    if (!map || !mapReady) return;
    const focusStore = mappedStores.find((store) => store.id === focusStoreId);
    if (focusStore) {
      setActiveStoreId(focusStore.id);
      map.setView([focusStore.latitude!, focusStore.longitude!], Math.min(Math.max(map.getZoom(), 15), 16), { animate: false });
      return;
    }
    setActiveStoreId("");
    if (userLocation) {
      map.setView([userLocation.latitude, userLocation.longitude], Math.min(Math.max(map.getZoom(), 15), 16), { animate: false });
    }
  }, [mapReady, focusStoreId, mappedStores, userLocation]);

  useEffect(() => {
    const map = leafletMap.current;
    if (!map || !mapReady || !userLocation || !locationFocusTick) return;
    setActiveStoreId("");
    map.setView([userLocation.latitude, userLocation.longitude], Math.min(Math.max(map.getZoom(), 15), 16), { animate: true });
  }, [mapReady, userLocation, locationFocusTick]);

  return (
    <div className="map-page">
      <section className="map-summary">
        <span>담당매장 {stores.length.toLocaleString()}개 · 완료 {completedCount.toLocaleString()}개 · 미완료 {(stores.length - completedCount).toLocaleString()}개</span>
        <span>{locationText}</span>
      </section>
      <section className="map-panel">
        <div ref={mapNode} className="store-map" />
      </section>
      <section className="panel map-active-panel">
        {activeStore ? (
          <>
            <div className="map-active-head">
              <div>
                {(() => {
                  const stats = statsByStore.get(activeStore.id) ?? emptyStats;
                  return <span className="map-active-stat">물품 {stats.completed.toLocaleString()}/{stats.total.toLocaleString()} · {isStoreComplete(activeStore, stats) ? "완료" : "미완료"}</span>;
                })()}
                <h2 title={activeStore.storeName}>{activeStore.storeName}</h2>
              </div>
              <StoreMoreMenu store={activeStore} onAssignToggle={() => onToggle(activeStore)} />
            </div>
            <p>{activeStore.storeAddress || "주소 없음"}</p>
            <div className="map-active-actions">
              <button type="button" onClick={() => onContacts(activeStore)}>담당자 정보</button>
              <button type="button" className="primary" onClick={() => onOpen(activeStore)}>입력</button>
            </div>
          </>
        ) : (
          <p className="muted">좌표가 있는 매장을 선택하면 여기에 정보가 표시됩니다.</p>
        )}
      </section>
    </div>
  );
}

function StoreAssignmentPanel({ stores, totalStores, statsByStore, geocoding, geocodeMessage, onGeocodeMissing, onGeocodeAll, onAssign, onAssignAll, onSave }: { stores: SurveyStore[]; totalStores: number; statsByStore: Map<string, RegionStats>; geocoding: boolean; geocodeMessage: string; onGeocodeMissing: () => void | Promise<void>; onGeocodeAll: () => void | Promise<void>; onAssign: (store: SurveyStore, assigned: boolean) => void | Promise<void>; onAssignAll: (stores: SurveyStore[], assigned: boolean) => void | Promise<void>; onSave: () => void | Promise<void> }) {
  const assignedCount = stores.filter((store) => store.mapIncluded === true).length;
  const missingCoordinateCount = stores.filter((store) => store.mapIncluded === true && !hasStoreCoordinates(store)).length;
  return (
    <section className="panel assignment-panel">
      <div className="assignment-head">
        <div>
          <h2>담당매장 관리</h2>
          <p>체크한 매장만 매장리스트와 매장지도에 표시됩니다. 위치정보를 가져오면 지도와 거리순을 사용할 수 있습니다.</p>
        </div>
        <strong>{assignedCount.toLocaleString()}<small>/{stores.length.toLocaleString()}</small></strong>
      </div>
      <div className="assignment-section">
        <div className="assignment-section-head">
          <strong>매장 위치정보</strong>
          <span>위치 없는 담당매장 {missingCoordinateCount.toLocaleString()}개</span>
        </div>
        <div className="assignment-location-actions">
          <button type="button" onClick={onGeocodeMissing} disabled={geocoding || missingCoordinateCount === 0}>{geocoding ? "검색 중" : "누락 위치만 가져오기"}</button>
          <button type="button" onClick={onGeocodeAll} disabled={geocoding || assignedCount === 0}>전체 위치 가져오기</button>
        </div>
      </div>
      {geocodeMessage && <p className="map-location-message">{geocodeMessage}</p>}
      <div className="assignment-list-head">
        <div className="assignment-list-title">
          <strong>매장 선택 ({assignedCount.toLocaleString()}/{totalStores.toLocaleString()})</strong>
          <span>* 위치없는 매장 {missingCoordinateCount.toLocaleString()}개</span>
        </div>
        <div className="assignment-actions">
          <button type="button" onClick={() => onAssignAll(stores, true)}>전체 선택</button>
          <button type="button" onClick={() => onAssignAll(stores, false)}>전체 해제</button>
        </div>
      </div>
      <div className="assignment-list">
        {stores.map((store, index) => {
          const stats = statsByStore.get(store.id) ?? emptyStats;
          return (
            <label key={store.id} className={`assignment-row ${store.mapIncluded === true ? "selected" : ""}`}>
              <input type="checkbox" checked={store.mapIncluded === true} onChange={(event) => onAssign(store, event.target.checked)} />
              <span className="assignment-order">{index + 1}</span>
              <span className="assignment-name" title={store.storeName}>{store.storeName}</span>
              <span className="assignment-address">{store.storeAddress || "주소 없음"}</span>
              <span className="assignment-stat">{stats.completed.toLocaleString()}/{stats.total.toLocaleString()}</span>
            </label>
          );
        })}
        {!stores.length && <p className="muted">검색 결과가 없습니다.</p>}
      </div>
      <button type="button" className="primary assignment-save" onClick={onSave}>저장하고 메인으로</button>
    </section>
  );
}

function ContactModal({ store, items, onClose }: { store?: SurveyStore; items: SurveyItem[]; onClose: () => void }) {
  if (!store) return null;
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal">
        <div className="modal-head">
          <div>
            <h2>담당자 리스트</h2>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        <Contacts items={items} />
      </section>
    </div>
  );
}

function SummaryModal({ region, stats, storeCount, completedStoreCount, mode, onClose }: { region?: string; stats: RegionStats; storeCount: number; completedStoreCount: number; mode: "regions" | "workspace" | "items"; onClose: () => void }) {
  const regionPercent = storeCount ? Math.round((completedStoreCount / storeCount) * 100) : 0;
  const itemPercent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;
  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <section className="modal">
        <div className="modal-head">
          <div>
            <h2>{region ?? "현재 지역"} 현황</h2>
            <p>{mode === "regions" ? "지역 기준 완료 현황입니다." : mode === "workspace" ? "매장과 물품 기준 진행률입니다." : "물품 기준 진행률입니다."}</p>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="닫기"><X size={18} /></button>
        </div>
        {(mode === "regions" || mode === "workspace") && (
          <div className="summary-progress-card">
            <div><span>{mode === "regions" ? "지역" : "매장"}</span><strong>{completedStoreCount.toLocaleString()}<small>/{storeCount.toLocaleString()}</small></strong></div>
            <div className="progress-line"><span style={{ width: `${regionPercent}%` }} /></div>
            <em>미완료 {(storeCount - completedStoreCount).toLocaleString()}</em>
          </div>
        )}
        {(mode === "workspace" || mode === "items") && (
          <div className="summary-progress-card">
            <div><span>물품</span><strong>{stats.completed.toLocaleString()}<small>/{stats.total.toLocaleString()}</small></strong></div>
            <div className="progress-line"><span style={{ width: `${itemPercent}%` }} /></div>
            <em>미완료 {(stats.total - stats.completed).toLocaleString()}</em>
          </div>
        )}
      </section>
    </div>
  );
}

function Contacts({ items }: { items: SurveyItem[] }) {
  const contacts = Array.from(new Map(items.map((item) => [`${item.companyManager}|${item.companyTel}|${item.martTel}`, item])).values());
  return (
    <section className="panel">
      <h2>담당자 정보</h2>
      {contacts.length === 0 && <p className="warn">확인 필요: 연락처 정보가 없습니다.</p>}
      {contacts.map((item) => {
        return (
          <div className="contact" key={`${item.companyManager}-${item.companyTel}-${item.martTel}`}>
            <dl className="contact-info">
              <dt>이름</dt><dd>{item.companyManager || "확인 필요"}</dd>
              <dt>연락처</dt><dd>{item.companyTel ? <a href={`tel:${item.companyTel.replace(/[^\d+]/g, "")}`}><Phone size={15} />{item.companyTel}</a> : <span className="warn">확인 필요</span>}</dd>
            </dl>
          </div>
        );
      })}
    </section>
  );
}

function ItemContact({ item }: { item: SurveyItem }) {
  const hasAnyContact = Boolean(item.companyManager || item.companyTel);
  return (
    <section className={`item-contact ${hasAnyContact && item.companyTel ? "" : "needs-check"}`}>
      <div>
        <h2>담당자 정보</h2>
        <span>이름: {item.companyManager || "확인 필요"}</span>
        <span>연락처: {item.companyTel ? <a href={`tel:${item.companyTel.replace(/[^\d+]/g, "")}`}>{item.companyTel}</a> : "확인 필요"}</span>
      </div>
    </section>
  );
}

function PhotoInput({ id, label, onFile }: { id?: string; label: string; onFile: (file: File) => void | Promise<void> }) {
  const pickId = `${id ?? uid("photo_pick")}-pick`;
  const cameraId = `${id ?? uid("photo_camera")}-camera`;
  return (
    <div className="photo-picker">
      <span>{label}</span>
      <div>
        <label className="photo-button" htmlFor={cameraId}><Camera size={18} />촬영<input id={cameraId} type="file" accept="image/*" capture="environment" onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])} /></label>
        <label className="photo-button" htmlFor={pickId}><Upload size={18} />선택<input id={pickId} type="file" accept="image/*" onChange={(event) => event.target.files?.[0] && onFile(event.target.files[0])} /></label>
      </div>
    </div>
  );
}

function ItemEditor({ item, storeItems, storeOperatingStatus, photos, fromBarcodeFlow, onSave, onSaved, onList, onStoreList, onMove, onBarcodeReturn, askConfirm }: { item: SurveyItem; storeItems: SurveyItem[]; storeOperatingStatus: StoreOperatingStatus | ""; photos: SurveyPhoto[]; fromBarcodeFlow?: boolean; onSave: (item: SurveyItem, photoOverride?: SurveyPhoto[]) => Promise<boolean>; onSaved: () => Promise<void>; onList: (focusId?: string) => void; onStoreList: () => void; onMove: (id: string) => void; onBarcodeReturn?: (id: string) => void; askConfirm: (options: ConfirmState) => Promise<boolean> }) {
  const [draft, setDraft] = useState(item);
  const [localPhotos, setLocalPhotos] = useState<SurveyPhoto[]>(photos);
  const [deletedPhotoIds, setDeletedPhotoIds] = useState<string[]>([]);
  const [photoMessage, setPhotoMessage] = useState("");
  const [priceOcrMessage, setPriceOcrMessage] = useState("");
  const [priceCandidates, setPriceCandidates] = useState<PriceCandidate[]>([]);
  const [saveMessage, setSaveMessage] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  useEffect(() => {
    setDraft(item);
    setLocalPhotos(photos);
    setDeletedPhotoIds([]);
    setPhotoMessage("");
    setPriceOcrMessage("");
    setPriceCandidates([]);
  }, [item.id]);
  useEffect(() => {
    if (!saveMessage) return;
    const timer = window.setTimeout(() => setSaveMessage(""), 5000);
    return () => window.clearTimeout(timer);
  }, [saveMessage]);
  useEffect(() => {
    if (!fromBarcodeFlow) return;
    const timer = window.setTimeout(() => {
      document.getElementById("item-status-section")?.scrollIntoView({ block: "start", behavior: "smooth" });
    }, 160);
    return () => window.clearTimeout(timer);
  }, [fromBarcodeFlow, item.id]);
  const update = (patch: Partial<SurveyItem>) => setDraft((old) => ({ ...old, ...patch, status: old.status === "미조사" ? "조사중" : old.status }));
  const missing = draft.normalDisplay ? requiredPhotoLabels(draft, localPhotos) : [];
  const itemPhotos = {
    display: localPhotos.find((photo) => photo.itemId === draft.id && photo.type === "PRODUCT_DISPLAY"),
    info: localPhotos.find((photo) => photo.itemId === draft.id && photo.type === "PRODUCT_INFO_BARCODE"),
    pos: localPhotos.find((photo) => photo.itemId === draft.id && photo.type === "POS_RECEIPT"),
  };
  const photoStateKey = (list: SurveyPhoto[]) => list.filter((photo) => photo.itemId === item.id).map((photo) => `${photo.id}:${photo.type}`).sort().join("|");
  const isDirty = useMemo(() => JSON.stringify({ ...draft, updatedAt: item.updatedAt }) !== JSON.stringify(item) || photoStateKey(localPhotos) !== photoStateKey(photos) || deletedPhotoIds.length > 0, [draft, item, localPhotos, photos, deletedPhotoIds]);
  const confirmLeaveIfDirty = async () => {
    if (!isDirty) return true;
    return askConfirm({
      title: "저장하지 않고 이동할까요?",
      message: "해당 품목에서 입력한 내용은 저장되지 않고 이전 저장 상태로 돌아갑니다.",
      confirmText: "이동",
      cancelText: "계속 입력",
      danger: true,
    });
  };
  const runPriceOcr = async (blob: Blob) => {
    setPriceCandidates([]);
    setPriceOcrMessage("가격 인식 중...");
    try {
      const candidates = await detectPriceCandidatesFromBlob(blob);
      setPriceCandidates(candidates);
      setPriceOcrMessage(candidates.length > 0 ? "가격이 인식되었습니다." : "가격 인식 실패");
    } catch (error) {
      console.error(error);
      setPriceCandidates([]);
      setPriceOcrMessage("가격 인식 실패");
    }
  };
  const upload = async (type: PhotoType, file: File) => {
    const resized = await resizePhoto(file);
    const oldPhotos = localPhotos.filter((photo) => photo.itemId === draft.id && photo.type === type);
    setDeletedPhotoIds((old) => [...old, ...oldPhotos.filter((photo) => !photo.id.startsWith("temp_")).map((photo) => photo.id)]);
    const photo: SurveyPhoto = { id: uid("temp_photo"), region: draft.region, storeId: draft.storeId, itemId: draft.id, type, blob: resized.blob, originalName: file.name, mimeType: resized.mimeType, takenAt: now() };
    setLocalPhotos((old) => [...old.filter((candidate) => !(candidate.itemId === draft.id && candidate.type === type)), photo]);
    if (type === "PRODUCT_DISPLAY" || type === "POS_RECEIPT") {
      setPhotoMessage("");
      void runPriceOcr(resized.blob);
      return;
    }
    if (type !== "PRODUCT_INFO_BARCODE") {
      setPhotoMessage("");
      return;
    }

    try {
      const detected = await detectBarcodeFromFile(file);
      if (!detected.supported) {
        setPhotoMessage("이 브라우저는 바코드 자동인식을 지원하지 않습니다.");
        return;
      }
      const expected = onlyDigits(draft.barcode);
      const detectedValues = detected.values.map(onlyDigits).filter(Boolean);
      if (detectedValues.length === 0) {
        setPhotoMessage("바코드를 인식하지 못했습니다.");
        return;
      }
      const matched = expected ? detectedValues.includes(expected) : false;
      const detectedText = detectedValues.join(", ");
      if (expected && matched) {
        update({ barcodeMatch: "O" });
        setPhotoMessage(`바코드 일치: ${expected}`);
      } else if (expected) {
        update({ barcodeMatch: "X" });
        setPhotoMessage(`바코드 불일치: 조사표 ${expected} / 촬영 ${detectedText}`);
      } else {
        setPhotoMessage(`바코드 인식: ${detectedText}`);
      }
    } catch (error) {
      console.error(error);
      setPhotoMessage("바코드 자동인식에 실패했습니다.");
    }
  };
  const nextTodoId = () => {
    const currentIndex = storeItems.findIndex((candidate) => candidate.id === draft.id);
    const ordered = [...storeItems.slice(currentIndex + 1), ...storeItems.slice(0, Math.max(0, currentIndex))];
    return ordered.find((candidate) => candidate.id !== draft.id && candidate.status !== "완료")?.id;
  };
  const nextSequentialId = () => {
    const currentIndex = storeItems.findIndex((candidate) => candidate.id === draft.id);
    return storeItems[currentIndex + 1]?.id;
  };
  const appendMemo = (text: string) => {
    const parts = draft.memo.split("/").map((part) => part.trim()).filter(Boolean);
    if (parts.includes(text)) return draft.memo;
    return parts.length ? `${parts.join(" / ")} / ${text}` : text;
  };
  const toggleMemo = (text: string) => {
    const parts = draft.memo.split("/").map((part) => part.trim()).filter(Boolean);
    return parts.includes(text) ? parts.filter((part) => part !== text).join(" / ") : appendMemoText(draft.memo, text);
  };
  const removeLocalPhotoTypes = (types: PhotoType[]) => {
    const removing = localPhotos.filter((photo) => photo.itemId === draft.id && types.includes(photo.type));
    if (removing.length === 0) return;
    setDeletedPhotoIds((old) => [...old, ...removing.filter((photo) => !photo.id.startsWith("temp_")).map((photo) => photo.id)]);
    setLocalPhotos((old) => old.filter((photo) => !(photo.itemId === draft.id && types.includes(photo.type))));
  };
  const updateNormalDisplay = async (value: string) => {
    const nextValue = value as SurveyItem["normalDisplay"];
    const removeTypes = nextValue === "X" ? ["PRODUCT_DISPLAY", "PRODUCT_INFO_BARCODE"] as PhotoType[] : [];
    const removing = localPhotos.filter((photo) => photo.itemId === draft.id && removeTypes.includes(photo.type));
    if (removing.length > 0) {
      const ok = await askConfirm({
        title: "",
        message: `진열여부를 변경하면 현재 선택한 상태에 맞지 않는 사진 ${removing.length}장이 삭제됩니다. 계속할까요?`,
        confirmText: "삭제 후 변경",
        cancelText: "취소",
        danger: true,
        plain: true,
      });
      if (!ok) return;
      removeLocalPhotoTypes(removeTypes);
    }
    update({
      normalDisplay: nextValue,
      photoCase: nextValue === "X" ? "POS_ONLY" : nextValue === "O" ? "NORMAL" : "",
      specMatch: nextValue === "X" ? "" : draft.specMatch,
      barcodeMatch: nextValue === "X" ? "" : draft.barcodeMatch,
      barcodeRegistered: nextValue === "O" ? "" : draft.barcodeRegistered,
      abnormalStatus: nextValue === "O" ? "" : draft.abnormalStatus,
      posChecked: nextValue === "O" ? "" : draft.posChecked,
      posPrice: null,
      abnormalDisplay: nextValue === "X" ? "" : draft.abnormalDisplay,
      memo: nextValue === "O" ? removeMemoTexts(draft.memo, POS_MEMOS) : draft.memo,
    });
  };
  const cleanDiscountMemo = () => removeMemoTexts(draft.memo, ["상시할인", "할인 정보 확인 불가"]);
  const updatePosChecked = (value: string) => {
    const cleanMemo = removeMemoTexts(draft.memo, POS_MEMOS);
    const memo = value === "조회함" ? appendMemoText(cleanMemo, "POS 조회") : value === "조회불가" ? appendMemoText(cleanMemo, "POS 조회 불가") : cleanMemo;
    update({
      posChecked: value as SurveyItem["posChecked"],
      memo,
    });
  };
  const updateDiscountEnabled = (value: string) => {
    const hasDiscount = value === "할인 있음";
    update({
      hasDiscount,
      discountPrice: hasDiscount ? draft.discountPrice : null,
      discountStartDate: hasDiscount ? draft.discountStartDate : "",
      discountEndDate: hasDiscount ? draft.discountEndDate : "",
      discountType: hasDiscount ? draft.discountType : "",
      discountOral: hasDiscount ? draft.discountOral : false,
      discountPeriodMode: hasDiscount ? draft.discountPeriodMode : "",
      memo: hasDiscount ? draft.memo : removeMemoTexts(draft.memo, ["1+1 행사", "상시할인", "할인 정보 확인 불가", "구두확인"]),
    });
  };
  const updateDiscountMode = (mode: NonNullable<SurveyItem["discountPeriodMode"]>) => {
    const baseMemo = cleanDiscountMemo();
    if (mode === "상시할인") {
      update({
        discountPeriodMode: mode,
        discountStartDate: "",
        discountEndDate: "",
        discountType: "②",
        memo: appendMemoText(baseMemo, "상시할인"),
      });
      return;
    }
    if (mode === "모름") {
      update({
        discountPeriodMode: mode,
        discountStartDate: "",
        discountEndDate: "",
        discountType: "",
        memo: appendMemoText(baseMemo, "할인 정보 확인 불가"),
      });
      return;
    }
    update({
      discountPeriodMode: mode,
      discountType: periodTypeFromDates(draft.discountStartDate, draft.discountEndDate),
      memo: baseMemo,
    });
  };
  const updateOnePlusOne = (checked: boolean) => {
    const memo = checked ? appendMemo("1+1 행사") : removeMemoTexts(draft.memo, ["1+1 행사"]);
    update({
      hasDiscount: checked ? true : draft.hasDiscount,
      discountPrice: checked && draft.normalPrice !== null ? Math.round(draft.normalPrice / 2) : null,
      memo,
    });
  };
  const updateDiscountDate = (field: "discountStartDate" | "discountEndDate", value: string) => {
    const nextStart = field === "discountStartDate" ? value : draft.discountStartDate;
    const nextEnd = field === "discountEndDate" ? value : draft.discountEndDate;
    update({
      [field]: value,
      discountPeriodMode: "기간 할인",
      discountType: periodTypeFromDates(nextStart, nextEnd),
    });
  };
  const priceBlocked = draft.normalDisplay === "X" && draft.barcodeRegistered === "X";
  const priceFeedback = draft.basePrice !== null && draft.normalPrice !== null
    ? (() => {
        const diff = draft.normalPrice - draft.basePrice!;
        const percent = draft.basePrice ? Math.round((Math.abs(diff) / draft.basePrice) * 100) : 0;
        const messages = [{ type: diff < 0 ? "warn" : "ok", text: diff < 0 ? "조사가격이 기준가격보다 작습니다." : diff > 0 ? "조사가격이 기준가격보다 큽니다." : "조사가격이 기준가격과 같습니다." }];
        if (percent >= PRICE_DIFF_WARN_PERCENT) messages.push({ type: "warn", text: `기준가격과 ${percent}% 차이납니다.` });
        return { messages };
      })()
    : undefined;
  const storeSaveLocked = storeOperatingStatus !== "영업 중";
  const handleSave = async () => {
    if (storeSaveLocked) {
      setSaveMessage("매장 상태가 영업 중일 때만 저장할 수 있습니다.");
      return;
    }
    setIsSaving(true);
    setSaveMessage("저장 중...");
    try {
      const saved = await onSave(draft, localPhotos);
      if (saved) {
        await Promise.all(deletedPhotoIds.map((id) => deletePhoto(id)));
        const persistedPhotos = await Promise.all(localPhotos.filter((photo) => photo.id.startsWith("temp_")).map(async (photo) => {
          const persisted = { ...photo, id: uid("photo") };
          await putPhoto(persisted);
          return persisted;
        }));
        setLocalPhotos((old) => [...old.filter((photo) => !photo.id.startsWith("temp_")), ...persistedPhotos]);
        setDeletedPhotoIds([]);
        await onSaved();
        setDraft((old) => ({ ...old, status: "완료" }));
        setSaveMessage(`저장 완료 · ${new Date().toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" })}`);
        if (fromBarcodeFlow && onBarcodeReturn) {
          if (await askConfirm({ title: "저장되었습니다", message: "바코드 창으로 돌아갈까요?", confirmText: "이동", cancelText: "현재 품목 보기" })) onBarcodeReturn(draft.id);
          return;
        }
        const nextId = nextTodoId();
        if (nextId) {
          if (await askConfirm({ title: "저장되었습니다", message: "다음 미등록 상품으로 이동할까요?", confirmText: "이동", cancelText: "현재 품목 보기" })) onMove(nextId);
        } else if (await askConfirm({ title: "전 품목 입력완료", message: "매장리스트로 돌아갈까요?", confirmText: "예", cancelText: "아니오" })) {
          onStoreList();
        }
      } else {
        setSaveMessage("저장이 취소되었습니다.");
      }
    } catch (error) {
      console.error(error);
      setSaveMessage("저장 실패: 다시 눌러주세요.");
    } finally {
      setIsSaving(false);
    }
  };
  const goListWithoutSave = async () => {
    if (await confirmLeaveIfDirty()) {
      if (fromBarcodeFlow && onBarcodeReturn) onBarcodeReturn(draft.id);
      else onList(draft.id);
    }
  };
  const saveAndNext = async () => {
    const nextId = nextSequentialId();
    if (nextId) {
      if (await confirmLeaveIfDirty()) onMove(nextId);
    }
    else setSaveMessage("마지막 품목입니다.");
  };
  return <main className="page item-page"><section className="item-hero"><div className="item-hero-row"><span className="item-code">{draft.itemNo}</span><strong className="item-hero-name" title={draft.productName}>{draft.productName}</strong><Badge text={draft.status} /></div></section>
    <ItemContact item={draft} />
    <details className="panel" open><summary>① 국군복지단 제시정보</summary><Info item={draft} /></details>
    <section className="panel"><h2>② 실물 확인</h2><Choice label="진열여부" note="*조사표에는 정상진열로 표기" value={draft.normalDisplay} values={["O", "X"]} onChange={updateNormalDisplay} /><Choice label="규격일치" disabled={draft.normalDisplay !== "O"} value={draft.normalDisplay === "O" ? draft.specMatch : ""} values={["O", "X"]} onChange={(value) => update({ specMatch: value as SurveyItem["specMatch"] })} /><Choice label="바코드일치" disabled={draft.normalDisplay !== "O"} value={draft.normalDisplay === "O" ? draft.barcodeMatch : ""} values={["O", "X"]} onChange={(value) => update({ barcodeMatch: value as SurveyItem["barcodeMatch"] })} /></section>
    <section id="item-status-section" className={`panel ${draft.normalDisplay === "X" ? "" : "disabled-block"}`}><h2>③ 상태 <small className="section-note">(정상진열 X 시 입력)</small></h2><Choice label="바코드 등록 여부" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.barcodeRegistered : ""} values={["O", "X"]} onChange={(value) => update({ barcodeRegistered: value as SurveyItem["barcodeRegistered"] })} /><Choice label="판매여부" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.abnormalStatus : ""} values={["미진열", "미판매"]} onChange={(value) => update({ abnormalStatus: value as SurveyItem["abnormalStatus"] })} /><Choice label="POS 조회 여부" disabled={draft.normalDisplay !== "X"} value={draft.normalDisplay === "X" ? draft.posChecked : ""} values={["조회함", "조회불가"]} onChange={updatePosChecked} /></section>
    <section className="panel">
      <h2 className="section-title-row">④ 사진자료 {missing.length > 0 && <span className="inline-missing">사진누락: {missing.join(", ")}</span>}</h2>
      {!draft.normalDisplay && <p className="notice">먼저 ② 실물 확인에서 진열여부 O/X를 선택해 주세요.</p>}
      {draft.normalDisplay === "O" && <p className="small-help barcode-help">참고: 제품정보사진 촬영 시 브라우저가 지원하면 바코드를 자동 비교합니다.</p>}
      <PhotoSlot id="photo-product-display" label="제품진열사진" description="가격정보와 진열상품이 동시노출 되도록 촬영" disabled={draft.normalDisplay !== "O"} photo={itemPhotos.display} message={priceOcrMessage} messageTone={priceCandidates.length ? "ok" : priceOcrMessage.includes("중") ? "pending" : "warn"} onFile={(file) => upload("PRODUCT_DISPLAY", file)} onDelete={(photo) => { setDeletedPhotoIds((old) => photo.id.startsWith("temp_") ? old : [...old, photo.id]); setLocalPhotos((old) => old.filter((candidate) => candidate.id !== photo.id)); }} />
      <PhotoSlot id="photo-product-info" label="제품정보사진" description="상품후면 제품상세정보와 바코드 동시노출 되도록 촬영" disabled={draft.normalDisplay !== "O"} photo={itemPhotos.info} message={photoMessage} messageTone={photoMessage.includes("불일치") || photoMessage.includes("실패") || photoMessage.includes("못했습니다") ? "warn" : "ok"} onFile={(file) => upload("PRODUCT_INFO_BARCODE", file)} onDelete={(photo) => { setDeletedPhotoIds((old) => photo.id.startsWith("temp_") ? old : [...old, photo.id]); setLocalPhotos((old) => old.filter((candidate) => candidate.id !== photo.id)); }} />
      <PhotoSlot id="photo-pos-receipt" label="POS/영수증사진" description="제품진열사진으로 가격정보 확인불가 시 POS기 또는 영수증 촬영" disabled={!draft.normalDisplay} photo={itemPhotos.pos} message={itemPhotos.pos ? priceOcrMessage : ""} messageTone={priceCandidates.length ? "ok" : priceOcrMessage.includes("중") ? "pending" : "warn"} onFile={(file) => upload("POS_RECEIPT", file)} onDelete={(photo) => { setDeletedPhotoIds((old) => photo.id.startsWith("temp_") ? old : [...old, photo.id]); setLocalPhotos((old) => old.filter((candidate) => candidate.id !== photo.id)); }} />
    </section>
    <section className={`panel price-panel ${priceBlocked ? "disabled-block" : ""}`}>
      <h2>⑤ 가격</h2>
      <p className="price-base">기준가격: <strong>{draft.basePrice?.toLocaleString() ?? "-"}원</strong></p>
      {priceBlocked && <p className="small-help warn">바코드 미등록 미판매 상품은 가격 입력을 생략합니다.</p>}
      <PriceCandidateChips label="정상가 후보" candidates={priceCandidates} disabled={priceBlocked} onPick={(value) => update({ normalPrice: value, discountPrice: draft.memo.includes("1+1 행사") ? Math.round(value / 2) : draft.discountPrice })} />
      <Money label="정상가" disabled={priceBlocked} value={draft.normalPrice} onChange={(value) => { const normalPrice = num(value); update({ normalPrice, discountPrice: draft.memo.includes("1+1 행사") && normalPrice !== null ? Math.round(normalPrice / 2) : draft.discountPrice }); }} />
      {priceFeedback && <div className="price-feedback">{priceFeedback.messages.map((message) => <span className={message.type} key={message.text}><i aria-hidden="true">{message.type === "warn" ? "!" : "✓"}</i>{message.text}</span>)}</div>}
      <Choice label="할인 여부" disabled={priceBlocked} value={draft.hasDiscount === null ? "" : draft.hasDiscount ? "할인 있음" : "할인 없음"} values={["할인 없음", "할인 있음"]} onChange={updateDiscountEnabled} />
      <div className={draft.hasDiscount === false || priceBlocked ? "disabled-block" : ""}>
        <div className="field-row">
          <span>행사여부</span>
          <label className="pretty-check"><input type="checkbox" disabled={draft.hasDiscount === false || priceBlocked} checked={draft.memo.includes("1+1 행사")} onChange={(event) => updateOnePlusOne(event.target.checked)} /><i />1+1 행사</label>
        </div>
        <PriceCandidateChips label="할인가 후보" candidates={priceCandidates} disabled={draft.hasDiscount === false || priceBlocked} onPick={(value) => update({ hasDiscount: true, discountPrice: value })} />
        <Money label="할인가" value={draft.discountPrice} disabled={draft.hasDiscount === false || priceBlocked} onChange={(value) => update({ discountPrice: num(value) })} />
        <DiscountControls
          disabled={draft.hasDiscount === false || priceBlocked}
          mode={draft.discountPeriodMode ?? ""}
          oral={draft.discountOral ?? draft.discountType.includes("구두")}
          start={draft.discountStartDate}
          end={draft.discountEndDate}
          periodType={draft.discountType}
          onMode={updateDiscountMode}
          onOral={(discountOral) => update({ discountOral, memo: discountOral ? appendMemo("구두확인") : removeMemoTexts(draft.memo, ["구두확인"]) })}
          onDate={updateDiscountDate}
        />
      </div>
    </section>
    <section className="panel"><h2>⑥ 특이사항</h2><div className={`abnormal-block ${draft.normalDisplay === "X" ? "disabled-block" : ""}`}><Choice label="비정상진열" disabled={draft.normalDisplay === "X"} value={draft.normalDisplay === "X" ? "" : draft.abnormalDisplay ?? ""} values={["O", "X"]} onChange={(value) => update({ abnormalDisplay: value as SurveyItem["abnormalDisplay"] })} />{draft.abnormalDisplay === "O" && draft.normalDisplay !== "X" && <p className="small-help warn">비정상진열이면 어떤 위치에 어떻게 진열되어 있었는지 아래 비고에 적어주세요.</p>}</div><div className="memo-block"><h3>비고</h3><p className="small-help">자주 쓰는 문구를 누르면 비고에 추가됩니다. 다시 누르면 해당 문구만 제거됩니다.</p><div className="chips memo-chips">{["가격 수기 입력", "폐점", "품절", "재고 소진", "재입고 예정", "1+1 행사", "임시휴업", "판매처 미협조"].map((text) => { const active = draft.memo.split("/").map((part) => part.trim()).includes(text); return <button key={text} className={active ? "active" : ""} onClick={() => update({ memo: toggleMemo(text) })}>{text}</button>; })}</div><textarea placeholder="예: 판매처 미협조 / 재입고 예정 / 사진 촬영 불가" value={draft.memo} onChange={(event) => update({ memo: event.target.value })} /></div></section>
    {saveMessage && <div className={`save-toast ${saveMessage.includes("실패") ? "danger-toast" : ""}`}>{saveMessage}</div>}
    <div className="item-action-fab">
      <div className="item-progress-mini"><span style={{ width: `${storeItems.length ? Math.round((storeItems.filter((candidate) => candidate.status === "완료").length + (draft.status === "완료" && !storeItems.find((candidate) => candidate.id === draft.id && candidate.status === "완료") ? 1 : 0)) / storeItems.length * 100) : 0}%` }} /></div>
      <button type="button" onClick={goListWithoutSave} disabled={isSaving}>목록</button>
      <button type="button" className="primary" onClick={handleSave} disabled={isSaving || storeSaveLocked} aria-label="저장" title={storeSaveLocked ? "매장 상태가 영업 중일 때만 저장할 수 있습니다." : undefined}><CheckCircle2 size={19} />{isSaving ? "저장 중" : "저장"}</button>
      <button type="button" onClick={saveAndNext} disabled={isSaving}>다음</button>
    </div>
  </main>;
}

function PriceCandidateChips({ label, candidates, disabled, onPick }: { label: string; candidates: PriceCandidate[]; disabled?: boolean; onPick: (value: number) => void }) {
  if (candidates.length === 0) return null;
  return (
    <div className={`price-candidates ${disabled ? "disabled-block" : ""}`}>
      <span>{label}</span>
      <div>
        {candidates.map((candidate) => (
          <button type="button" key={`${label}-${candidate.value}`} disabled={disabled} onClick={() => onPick(candidate.value)}>
            {candidate.value.toLocaleString()}원
          </button>
        ))}
      </div>
    </div>
  );
}

function PhotoSlot({ id, label, description, disabled, photo, message, messageTone, onFile, onDelete }: { id: string; label: string; description: string; disabled?: boolean; photo?: SurveyPhoto; message?: string; messageTone?: "ok" | "warn" | "pending"; onFile: (file: File) => void | Promise<void>; onDelete: (photo: SurveyPhoto) => void | Promise<void> }) {
  return (
    <div id={`${id}-slot`} className={`photo-slot ${photo ? "uploaded" : ""} ${disabled ? "photo-disabled" : ""}`}>
      <div>
        <div className="photo-title">
          <strong>{label}</strong>
          <small>{description}</small>
        </div>
      </div>
      {photo && <PhotoPreview photo={photo} className="wide-preview" />}
      <div className="photo-actions">
        {!photo && !disabled && <PhotoInput id={id} label="촬영/선택" onFile={onFile} />}
        {photo && !disabled && <button className="danger" onClick={() => onDelete(photo)}>지우기</button>}
        {disabled && <span className="photo-disabled-note">진열여부 선택에 따라 비활성화됨</span>}
      </div>
      {message && !disabled && <p className={`upload-message photo-result ${messageTone ?? "ok"}`}>{message}</p>}
    </div>
  );
}

function Info({ item }: { item: SurveyItem }) {
  return <dl className="info"><dt>물품코드</dt><dd>{item.itemNo}</dd><dt>상세주소</dt><dd>{item.detailAddress || "-"}</dd><dt>제조사</dt><dd>{item.companyName}</dd><dt>물품명</dt><dd>{item.productName}</dd><dt>규격</dt><dd>{item.spec}</dd><dt>기준가격</dt><dd>{item.basePrice !== null ? `${item.basePrice.toLocaleString()}원` : "-"}</dd><dt>바코드</dt><dd>{item.barcode}</dd></dl>;
}

function DiscountControls({
  disabled,
  mode,
  oral,
  start,
  end,
  periodType,
  onMode,
  onOral,
  onDate,
}: {
  disabled?: boolean;
  mode: NonNullable<SurveyItem["discountPeriodMode"]>;
  oral: boolean;
  start: string;
  end: string;
  periodType: string;
  onMode: (mode: NonNullable<SurveyItem["discountPeriodMode"]>) => void;
  onOral: (oral: boolean) => void;
  onDate: (field: "discountStartDate" | "discountEndDate", value: string) => void;
}) {
  const normalized = periodType.replace("구두", "");
  const datesDisabled = disabled || mode !== "기간 할인";
  return (
    <div className="discount-controls">
      <div className="field-row discount-period-row">
        <span>할인기간</span>
        <div className="period-control">
          <div className="segmented">
            {(["상시할인", "기간 할인", "모름"] as NonNullable<SurveyItem["discountPeriodMode"]>[]).map((candidate) => (
              <button disabled={disabled} className={mode === candidate ? "active" : ""} key={candidate} onClick={() => onMode(mode === candidate ? "" : candidate)}>{candidate}</button>
            ))}
          </div>
          <label className="pretty-check"><input type="checkbox" disabled={disabled} checked={oral} onChange={(event) => onOral(event.target.checked)} /><i />구두 확인</label>
          <div className={`date-range ${datesDisabled ? "range-disabled" : ""}`}>
            <input aria-label="할인 시작일" type="date" disabled={datesDisabled} value={start} onChange={(event) => onDate("discountStartDate", event.target.value)} />
            <b>~</b>
            <input aria-label="할인 종료일" type="date" disabled={datesDisabled} value={end} onChange={(event) => onDate("discountEndDate", event.target.value)} />
          </div>
          <div className="readonly-period">
            <span className={normalized === "①" ? "active" : ""}>① 31일 이내</span>
            <span className={normalized === "②" ? "active" : ""}>② 32일 이상</span>
            {!normalized && <em>{mode === "모름" ? "기간 정보 없음" : "날짜 입력 시 자동 확인"}</em>}
            {oral && normalized && <em>{normalized}구두확인</em>}
          </div>
        </div>
      </div>
    </div>
  );
}

function Choice({ label, note, value, values, disabled, onChange }: { label: string; note?: string; value: string; values: string[]; disabled?: boolean; onChange: (value: string) => void | Promise<void> }) {
  return <div className="field-row"><span>{label}{note && <small className="field-note">{note}</small>}</span><div className="segmented">{values.map((candidate) => <button disabled={disabled} className={value === candidate ? "active" : ""} key={candidate} onClick={() => onChange(value === candidate ? "" : candidate)}>{candidate}</button>)}</div></div>;
}

function Money({ label, value, disabled, onChange }: { label: string; value: number | null; disabled?: boolean; onChange: (value: string) => void }) {
  return <label>{label}<input inputMode="numeric" enterKeyHint="done" pattern="[0-9,]*" disabled={disabled} value={typeof value === "number" && Number.isFinite(value) ? value.toLocaleString() : ""} onChange={(event) => onChange(event.target.value.replace(/\D/g, ""))} onKeyDown={(event) => { if (event.key === "Enter") event.currentTarget.blur(); }} placeholder="원" /></label>;
}

function Validation({ title, items, open }: { title: string; items: SurveyItem[]; open: (id: string) => void }) {
  return <section className="panel"><h2>{title} ({items.length.toLocaleString()}개)</h2>{items.slice(0, 80).map((item) => <button className="row-button" key={item.id} onClick={() => open(item.id)}>{item.itemNo} · {item.productName} · {item.storeName}</button>)}</section>;
}

export default App;
