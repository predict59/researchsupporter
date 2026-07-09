import type { RegionStats, SurveyItem, SurveyPhoto, SurveyStore } from "./types";

export function photosForItem(item: SurveyItem, storePhotos: SurveyPhoto[]) {
  return storePhotos.filter((photo) => photo.storeId === item.storeId && (!photo.itemId || photo.itemId === item.id));
}

export function requiredPhotoLabels(item: SurveyItem, photos: SurveyPhoto[]) {
  const hasFront = photos.some((photo) => photo.type === "STORE_FRONT");
  const missing = productPhotoMissingLabels(item, photos);
  if (!hasFront) return ["매장사진", ...missing];
  return missing;
}

export function productPhotoMissingLabels(item: SurveyItem, photos: SurveyPhoto[]) {
  if (item.memo.includes("판매처 폐점") || item.memo.includes("임시휴업")) return [];
  const photoCase = photoCaseOf(item);
  const hasDisplay = photos.some((photo) => photo.type === "PRODUCT_DISPLAY" && photo.itemId === item.id);
  const hasInfo = photos.some((photo) => photo.type === "PRODUCT_INFO_BARCODE" && photo.itemId === item.id);
  const missing: string[] = [];
  if (photoCase === "MISSING") {
    missing.push("물품 사진 전체 누락");
  } else if (photoCase === "POS_ONLY") {
    return missing;
  } else {
    if (!hasDisplay) missing.push("제품진열사진");
    if (!hasInfo) missing.push("제품정보사진");
  }
  return missing;
}

export function photoCaseOf(item: SurveyItem) {
  if (item.normalDisplay === "X") return "POS_ONLY";
  if (item.normalDisplay === "O") return "NORMAL";
  return item.photoCase || "NORMAL";
}

export function isPhotoMissing(item: SurveyItem, photos: SurveyPhoto[]) {
  return requiredPhotoLabels(item, photos).length > 0;
}

export function summarize(items: SurveyItem[], photos: SurveyPhoto[]): RegionStats {
  return {
    total: items.length,
    completed: items.filter((item) => item.status === "완료").length,
    inProgress: items.filter((item) => item.status === "조사중").length,
    notStarted: items.filter((item) => item.status === "미조사").length,
    photoMissing: items.filter((item) => item.status === "완료" && productPhotoMissingLabels(item, photos).length > 0).length,
  };
}

export function storeStatus(store: SurveyStore, items: SurveyItem[]): SurveyStore["status"] {
  const own = items.filter((item) => item.storeId === store.id);
  if (own.length > 0 && own.every((item) => item.status === "완료")) return "완료";
  if (own.some((item) => item.status !== "미조사") || store.frontPhotoId) return "진행중";
  return "미시작";
}

export function safeFilePart(value: string) {
  return value.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_").slice(0, 80) || "미지정";
}

export function mapSearchAddress(address: string) {
  return address.replace(/\s+/g, " ").trim();
}

function isInAppBrowser() {
  return /KAKAOTALK|FBAN|FBAV|Instagram|Line\//i.test(navigator.userAgent);
}

function showInAppDownloadNotice(filename: string, openedFallback = false) {
  if (!isInAppBrowser()) return;
  window.setTimeout(() => {
    window.alert(`${filename}\n\n카카오톡 같은 인앱 브라우저에서는 파일 저장이 막힐 수 있습니다.${openedFallback ? "\n새로 열린 파일 화면에서 공유 또는 저장을 시도해 주세요." : "\n공유/저장 화면이 뜨지 않으면 다시 한 번 내려받기를 눌러 주세요."}`);
  }, 350);
}

export async function downloadBlob(blob: Blob, filename: string) {
  const file = new File([blob], filename, { type: blob.type || "application/octet-stream" });
  const shareTarget = navigator as Navigator & {
    canShare?: (data: ShareData) => boolean;
    share?: (data: ShareData) => Promise<void>;
  };
  if (shareTarget.canShare?.({ files: [file] }) && shareTarget.share) {
    try {
      await shareTarget.share({ files: [file], title: filename });
      return;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") return;
    }
  }
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.target = "_self";
  anchor.rel = "noopener";
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  const openedFallback = isInAppBrowser() ? Boolean(window.open(url, "_blank", "noopener,noreferrer")) : false;
  showInAppDownloadNotice(filename, openedFallback);
  setTimeout(() => {
    anchor.remove();
    URL.revokeObjectURL(url);
  }, isInAppBrowser() ? 180000 : 30000);
}
