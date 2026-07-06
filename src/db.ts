import { openDB } from "idb";
import type { AppSettings, Region, SurveyFileRecord, SurveyItem, SurveyPhoto, SurveyStore } from "./types";

const DB_NAME = "mw-price-survey-pwa";
const DB_VERSION = 1;

export const today = () => new Date().toISOString().slice(0, 10);
export const now = () => new Date().toISOString();
export const uid = (prefix: string) => `${prefix}_${crypto.randomUUID()}`;

const dbPromise = openDB(DB_NAME, DB_VERSION, {
  upgrade(db) {
    db.createObjectStore("surveyFiles", { keyPath: "id" });
    db.createObjectStore("regions", { keyPath: "name" });
    const stores = db.createObjectStore("stores", { keyPath: "id" });
    stores.createIndex("region", "region");
    const items = db.createObjectStore("items", { keyPath: "id" });
    items.createIndex("region", "region");
    items.createIndex("storeId", "storeId");
    const photos = db.createObjectStore("photos", { keyPath: "id" });
    photos.createIndex("region", "region");
    photos.createIndex("storeId", "storeId");
    photos.createIndex("itemId", "itemId");
    db.createObjectStore("settings", { keyPath: "id" });
  },
});

export async function getSettings(): Promise<AppSettings> {
  const db = await dbPromise;
  const saved = await db.get("settings", "app");
  return saved?.value ?? { defaultSurveyDate: today() };
}

export async function saveSettings(settings: AppSettings) {
  const db = await dbPromise;
  await db.put("settings", { id: "app", value: settings });
}

export async function getRegions() {
  return (await (await dbPromise).getAll("regions")) as Region[];
}

export async function saveParsedData(regions: Region[], stores: SurveyStore[], items: SurveyItem[]) {
  const db = await dbPromise;
  const tx = db.transaction(["regions", "stores", "items", "settings"], "readwrite");
  await Promise.all([
    ...regions.map((region) => tx.objectStore("regions").put(region)),
    ...stores.map((store) => tx.objectStore("stores").put(store)),
    ...items.map((item) => tx.objectStore("items").put(item)),
  ]);
  const settings = await tx.objectStore("settings").get("app");
  await tx.objectStore("settings").put({
    id: "app",
    value: { defaultSurveyDate: today(), ...settings?.value, currentRegion: settings?.value?.currentRegion },
  });
  await tx.done;
}

export async function getBarcodeIndex() {
  const db = await dbPromise;
  const record = (await db.get("surveyFiles", "barcodeIndex")) as SurveyFileRecord | undefined;
  return record?.value ?? {};
}

export async function saveBarcodeIndex(value: Record<string, string>) {
  const db = await dbPromise;
  await db.put("surveyFiles", { id: "barcodeIndex", value, updatedAt: now() } satisfies SurveyFileRecord);
}

export async function getStores(region?: string) {
  const db = await dbPromise;
  if (!region) return (await db.getAll("stores")) as SurveyStore[];
  return (await db.getAllFromIndex("stores", "region", region)) as SurveyStore[];
}

export async function getItems(region?: string) {
  const db = await dbPromise;
  if (!region) return (await db.getAll("items")) as SurveyItem[];
  return (await db.getAllFromIndex("items", "region", region)) as SurveyItem[];
}

export async function getStoreItems(storeId: string) {
  return (await (await dbPromise).getAllFromIndex("items", "storeId", storeId)) as SurveyItem[];
}

export async function putStore(store: SurveyStore) {
  await (await dbPromise).put("stores", store);
}

export async function putItem(item: SurveyItem) {
  await (await dbPromise).put("items", item);
}

export async function getPhotosByRegion(region: string) {
  return (await (await dbPromise).getAllFromIndex("photos", "region", region)) as SurveyPhoto[];
}

export async function getPhotos() {
  return (await (await dbPromise).getAll("photos")) as SurveyPhoto[];
}

export async function getPhotosByStore(storeId: string) {
  return (await (await dbPromise).getAllFromIndex("photos", "storeId", storeId)) as SurveyPhoto[];
}

export async function putPhoto(photo: SurveyPhoto) {
  await (await dbPromise).put("photos", photo);
}

export async function deletePhoto(id: string) {
  await (await dbPromise).delete("photos", id);
}

export async function importRegionData(region: string, stores: SurveyStore[], items: SurveyItem[], photos: SurveyPhoto[]) {
  const db = await dbPromise;
  const tx = db.transaction(["regions", "stores", "items", "photos"], "readwrite");
  const oldStores = (await tx.objectStore("stores").index("region").getAll(region)) as SurveyStore[];
  const oldItems = (await tx.objectStore("items").index("region").getAll(region)) as SurveyItem[];
  const oldPhotos = (await tx.objectStore("photos").index("region").getAll(region)) as SurveyPhoto[];
  await Promise.all([
    ...oldStores.map((store) => tx.objectStore("stores").delete(store.id)),
    ...oldItems.map((item) => tx.objectStore("items").delete(item.id)),
    ...oldPhotos.map((photo) => tx.objectStore("photos").delete(photo.id)),
  ]);
  await tx.objectStore("regions").put({ name: region, updatedAt: now() });
  await Promise.all([
    ...stores.map((store) => tx.objectStore("stores").put(store)),
    ...items.map((item) => tx.objectStore("items").put(item)),
    ...photos.map((photo) => tx.objectStore("photos").put(photo)),
  ]);
  await tx.done;
}

export async function importAllData(regions: Region[], stores: SurveyStore[], items: SurveyItem[], photos: SurveyPhoto[], settings: AppSettings) {
  const db = await dbPromise;
  const names = ["surveyFiles", "regions", "stores", "items", "photos", "settings"];
  const tx = db.transaction(names, "readwrite");
  await Promise.all(names.map((name) => tx.objectStore(name).clear()));
  await Promise.all([
    ...regions.map((region) => tx.objectStore("regions").put(region)),
    ...stores.map((store) => tx.objectStore("stores").put(store)),
    ...items.map((item) => tx.objectStore("items").put(item)),
    ...photos.map((photo) => tx.objectStore("photos").put(photo)),
  ]);
  await tx.objectStore("settings").put({ id: "app", value: settings });
  await tx.done;
}

export async function clearAllData() {
  const db = await dbPromise;
  const names = ["surveyFiles", "regions", "stores", "items", "photos", "settings"];
  const tx = db.transaction(names, "readwrite");
  await Promise.all(names.map((name) => tx.objectStore(name).clear()));
  await tx.done;
}
