export type Status = "미조사" | "조사중" | "완료";
export type StoreStatus = "미시작" | "진행중" | "완료";
export type StoreOperatingStatus = "영업 중" | "폐업" | "임시휴업";
export type PhotoType = "STORE_FRONT" | "PRODUCT_DISPLAY" | "PRODUCT_INFO_BARCODE" | "POS_RECEIPT";
export type PhotoCase = "" | "NORMAL" | "POS_ONLY" | "MISSING";

export type Region = {
  name: string;
  department?: string;
  city?: string;
  areaSummary?: string;
  updatedAt: string;
};

export type SurveyStore = {
  id: string;
  region: string;
  department?: string;
  city?: string;
  storeName: string;
  storeAddress: string;
  visitOrder?: number;
  mapIncluded?: boolean;
  latitude?: number;
  longitude?: number;
  geocodedAt?: string;
  geocodeStatus?: "성공" | "실패" | "미시도";
  operatingStatus?: StoreOperatingStatus;
  itemCount: number;
  completedCount: number;
  surveyDate: string;
  status: StoreStatus;
  frontPhotoId?: string;
  startedAt?: string;
  completedAt?: string;
  updatedAt: string;
};

export type SurveyItem = {
  id: string;
  itemNo: string;
  region: string;
  department?: string;
  city?: string;
  storeId: string;
  storeName: string;
  storeAddress: string;
  detailAddress?: string;
  companyName: string;
  companyTel: string;
  companyManager?: string;
  martName: string;
  martTel?: string;
  barcode: string;
  productName: string;
  spec: string;
  basePrice: number | null;
  surveyDate: string;
  normalDisplay: "" | "O" | "X";
  specMatch: "" | "O" | "X" | "-";
  barcodeMatch: "" | "O" | "X" | "-";
  normalPrice: number | null;
  hasDiscount: boolean | null;
  discountPrice: number | null;
  discountStartDate: string;
  discountEndDate: string;
  discountType: string;
  discountOral?: boolean;
  discountPeriodMode?: "" | "상시할인" | "기간 할인" | "모름";
  priceJudgment: "" | "동일" | "고가" | "저가" | "확인필요";
  abnormalDisplay?: "" | "O" | "X";
  photoCase?: PhotoCase;
  barcodeRegistered: "" | "O" | "X";
  abnormalStatus: "" | "미진열" | "미판매";
  posChecked: "" | "조회함" | "조회불가" | "미조회";
  posPrice: number | null;
  memo: string;
  status: Status;
  updatedAt: string;
};

export type SurveyPhoto = {
  id: string;
  region: string;
  itemId?: string;
  storeId: string;
  type: PhotoType;
  blob: Blob;
  originalName: string;
  mimeType: string;
  takenAt: string;
};

export type AppSettings = {
  currentRegion?: string;
  lastOpenedStoreId?: string;
  lastOpenedItemId?: string;
  defaultSurveyDate: string;
};

export type SurveyFileRecord = {
  id: "barcodeIndex";
  value: Record<string, string>;
  updatedAt: string;
};

export type RegionStats = {
  total: number;
  completed: number;
  inProgress: number;
  notStarted: number;
  photoMissing: number;
};

export type BackupPayload = {
  version: 1;
  exportedAt: string;
  scope: "region" | "all";
  region?: string;
  regions: Region[];
  stores: SurveyStore[];
  items: SurveyItem[];
  photos: Array<Omit<SurveyPhoto, "blob"> & { dataUrl: string }>;
  settings: AppSettings;
};
