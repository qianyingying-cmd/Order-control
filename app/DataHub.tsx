"use client";

import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type DataKind = "sales" | "wholesale" | "inventory" | "oih";
type Row = Record<string, unknown>;
type DetailRecord = {
  month: string;
  sku: string;
  name: string;
  brand: string;
  country: string;
  category: string;
  middleCategory: string;
  smallCategory: string;
  series: string;
  channel: string;
  sourceType?: "零售" | "批发";
  quantity: number;
  amount: number;
  revenue: number;
};

type Dataset = {
  schemaVersion?: number;
  kind: DataKind;
  fileName: string;
  uploadedAt: string;
  rowCount: number;
  sheetName: string;
  headers: string[];
  status: "ready" | "mapping";
  message: string;
  summary: {
    quantity: number;
    amount?: number;
    revenue?: number;
    skuCount: number;
    monthCount: number;
    countryCount: number;
    latestMonth: string;
    detectedDateField?: string;
  };
  monthly: Record<string, number>;
  sku: Record<string, number>;
  countries: Record<string, number>;
  details?: DetailRecord[];
};

const DB_NAME = "wilson-rolling-inventory";
const STORE = "datasets";
const DATA_SCHEMA_VERSION = 9;
const SALES_SCHEMA_VERSION = 10;
const CANONICAL_BRANDS = ["ANTA", "FILA", "DESCENTE", "SALOMON", "WILSON", "ARC'TERYX"] as const;
const kinds: Array<{ kind: DataKind; title: string; subtitle: string; accent: string }> = [
  { kind: "sales", title: "零售销售", subtitle: "观远 R01 / 门店零售", accent: "blue" },
  { kind: "wholesale", title: "批发销售", subtitle: "观远批发 / 客户销售", accent: "blue" },
  { kind: "inventory", title: "月末库存", subtitle: "BI 库存余额 / M01", accent: "green" },
  { kind: "oih", title: "OIH / 订单", subtitle: "在途、已下单及 VBR8", accent: "amber" },
];

const aliases = {
  sku: ["sku", "货号", "商品编码", "物料编码", "款号", "货品编码"],
  country: ["BI合并公司", "国家", "市场", "区域", "country", "market"],
  name: ["商品名称", "商品名", "品名", "货品名称", "产品名称", "sku名称"],
  brand: ["渠道品牌描述", "品牌", "品牌名称", "brand"],
  category: ["大类", "一级品类", "商品大类", "categoryl1", "category1"],
  middleCategory: ["中类", "二级品类", "商品中类", "categoryl2", "category2"],
  smallCategory: ["小类", "三级品类", "商品小类", "categoryl3", "category3", "品类"],
  series: ["系列", "商品系列", "产品系列", "series"],
  channel: ["渠道", "销售渠道", "店铺", "门店", "客户", "channel", "store"],
  storeType: ["经营类型", "门店类型", "店铺类型", "经营性质", "storetype"],
  onlinePlatform: ["一级平台（线上）", "一级平台线上", "线上平台", "电商平台", "平台"],
  date: ["单据/出库/收货月份", "单据出库收货月份", "单据月份", "出库月份", "收货月份", "日历月份", "库存日期", "单据日期", "出库日期", "收货日期", "过账日期", "销售日期", "库存月份", "快照日期", "日期", "月份", "年月", "date", "month"],
  sales: ["本期零售数量", "零售数量", "销售数量", "销量", "净销售数量", "销售件数", "数量", "qty", "quantity"],
  wholesale: ["本期批发数量", "批发数量", "批发销售数量", "出库数量", "发货数量", "销售数量", "销量", "数量", "qty", "quantity"],
  inventory: ["本期_期末库存数量", "本期期末库存数量", "期末库存数量", "库存数量", "可用库存", "数量", "qty", "quantity"],
  oih: ["oih", "在途数量", "订单数量", "下单数量", "未到货数量", "open order", "数量", "qty", "quantity"],
  salesValue: ["本期吊牌金额-人民币", "本期吊牌金额人民币", "吊牌销售金额", "销售吊牌金额", "零售吊牌金额", "吊牌金额", "retailvalue"],
  wholesaleValue: ["本期批发吊牌金额-人民币", "批发吊牌金额", "批发销售吊牌金额", "出库吊牌金额", "发货吊牌金额", "本期吊牌金额-人民币", "吊牌金额", "销售金额"],
  salesRevenue: ["本期零售金额-人民币", "本期零售金额人民币", "零售金额-人民币", "实际销售金额", "实收金额", "流水金额", "零售金额"],
  wholesaleRevenue: ["本期批发金额-人民币", "批发金额-人民币", "批发销售金额", "出库金额-人民币", "销售金额-人民币", "实际销售金额", "流水金额", "销售金额"],
  inventoryValue: ["本期_期末库存人民币吊牌金额", "本期期末库存人民币吊牌金额", "库存人民币吊牌金额", "期末库存人民币吊牌金额", "库存吊牌金额", "期末库存吊牌金额", "stockvalue"],
  oihValue: ["oih吊牌金额", "订单吊牌金额", "在途吊牌金额", "吊牌金额", "订单金额"],
};

function isSalesKind(kind: DataKind) {
  return kind === "sales" || kind === "wholesale";
}

function schemaVersionFor(kind: DataKind) {
  return isSalesKind(kind) ? SALES_SCHEMA_VERSION : DATA_SCHEMA_VERSION;
}

function normalizeBrand(value: unknown) {
  const raw = String(value ?? "").trim();
  const normalized = raw.toUpperCase().replace(/[\s_\-·•]/g, "");
  if (/ARCTERYX|ARC'TERYX|始祖鸟/.test(normalized)) return "ARC'TERYX";
  if (/DESCENTE|迪桑特/.test(normalized)) return "DESCENTE";
  if (/SALOMON|萨洛蒙/.test(normalized)) return "SALOMON";
  if (/WILSON|威尔胜/.test(normalized)) return "WILSON";
  if (/FILA|斐乐/.test(normalized)) return "FILA";
  if (/ANTA|安踏/.test(normalized)) return "ANTA";
  return "";
}

function normalizeCountry(value: unknown) {
  const raw = String(value ?? "").trim();
  if (/^(马来|马来西亚|malaysia)$/i.test(raw)) return "马来西亚";
  return raw;
}

function clean(value: unknown) {
  return String(value ?? "").trim().toLowerCase().replace(/[\s_\-\/（）()]/g, "");
}

function findHeader(headers: string[], options: string[]) {
  const normalized = options.map(clean);
  const exact = headers.find((header) => normalized.some((option) => clean(header) === option));
  if (exact) return exact;
  const ordered = [...normalized].sort((a, b) => b.length - a.length);
  return headers.find((header) => ordered.some((option) => option.length >= 2 && clean(header).includes(option)));
}

function findQuantityHeader(headers: string[], options: string[]) {
  return findHeader(headers.filter((header) => !/日期|月份|金额|吊牌|价格|单价/.test(header)), options);
}

function numeric(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const raw = String(value ?? "").replace(/,/g, "").trim();
  const direct = Number(raw);
  if (Number.isFinite(direct)) return direct;
  const parsed = Number(raw.replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function monthOf(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`;
  }
  if (typeof value === "number" && Number.isInteger(value) && value >= 200001 && value <= 209912) {
    const text = String(value);
    return `${text.slice(0, 4)}-${text.slice(4, 6)}`;
  }
  if (typeof value === "number" && value > 30000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? `${parsed.y}-${String(parsed.m).padStart(2, "0")}` : "";
  }
  const text = String(value ?? "").trim();
  const dayMonthYear = text.match(/^\d{1,2}[\/\-](\d{1,2})[\/\-](20\d{2})(?:\s|$)/);
  if (dayMonthYear) {
    return `${dayMonthYear[2]}-${String(Number(dayMonthYear[1])).padStart(2, "0")}`;
  }
  const match = text.match(/(20\d{2})\D?(0?[1-9]|1[0-2])(?:\D|$)/);
  if (match) return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? "" : `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function add(map: Record<string, number>, key: string, value: number) {
  if (key) map[key] = (map[key] ?? 0) + value;
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE)) request.result.createObjectStore(STORE, { keyPath: "kind" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getAll(): Promise<Dataset[]> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
    request.onsuccess = () => resolve(request.result as Dataset[]);
    request.onerror = () => reject(request.error);
  });
}

async function putOne(dataset: Dataset) {
  const db = await openDb();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(STORE, "readwrite").objectStore(STORE).put(dataset);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

async function deleteOne(kind: DataKind) {
  const db = await openDb();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(STORE, "readwrite").objectStore(STORE).delete(kind);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function parseWorkbook(buffer: ArrayBuffer, kind: DataKind, fileName: string): Dataset {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  let best: { sheetName: string; rows: Row[]; headers: string[]; score: number } | null = null;

  workbook.SheetNames.forEach((sheetName) => {
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, defval: "" });
    for (let index = 0; index < Math.min(25, matrix.length); index += 1) {
      const headers = (matrix[index] ?? []).map((value) => String(value).trim());
      const quantityOptions = kind === "sales" ? aliases.sales : kind === "wholesale" ? aliases.wholesale : kind === "inventory" ? aliases.inventory : aliases.oih;
      const score = [aliases.sku, aliases.country, aliases.date, quantityOptions].filter((options) => findHeader(headers, options)).length;
      if (!best || score > best.score) {
        const rows = matrix.slice(index + 1).map((values) =>
          Object.fromEntries(headers.map((header, column) => [header || `字段${column + 1}`, values[column]])),
        );
        best = { sheetName, rows, headers: headers.filter(Boolean), score };
      }
    }
  });

  if (!best) throw new Error("工作簿中没有可读取的数据");
  const quantityOptions = kind === "sales" ? aliases.sales : kind === "wholesale" ? aliases.wholesale : kind === "inventory" ? aliases.inventory : aliases.oih;
  const quantityHeader = findQuantityHeader(best.headers, quantityOptions);
  const valueOptions = kind === "sales" ? aliases.salesValue : kind === "wholesale" ? aliases.wholesaleValue : kind === "inventory" ? aliases.inventoryValue : aliases.oihValue;
  const valueHeader = findHeader(best.headers, valueOptions);
  const revenueOptions = kind === "wholesale" ? aliases.wholesaleRevenue : aliases.salesRevenue;
  const revenueHeader = isSalesKind(kind) ? findHeader(best.headers, revenueOptions) : undefined;
  const skuHeader = findHeader(best.headers, aliases.sku);
  const countryHeader = findHeader(best.headers, aliases.country);
  const dateHeader = findHeader(best.headers, aliases.date);
  const nameHeader = findHeader(best.headers, aliases.name);
  const brandHeader = findHeader(best.headers, aliases.brand);
  const categoryHeader = findHeader(best.headers, aliases.category);
  const middleCategoryHeader = findHeader(best.headers, aliases.middleCategory);
  const smallCategoryHeader = findHeader(best.headers, aliases.smallCategory);
  const seriesHeader = findHeader(best.headers, aliases.series);
  const channelHeader = findHeader(best.headers, aliases.channel);
  const storeTypeHeader = findHeader(best.headers, aliases.storeType);
  const onlinePlatformHeader = findHeader(best.headers, aliases.onlinePlatform);
  const monthly: Record<string, number> = {};
  const sku: Record<string, number> = {};
  const countries: Record<string, number> = {};
  const detailMap = new Map<string, DetailRecord>();
  let quantity = 0;
  let amount = 0;
  let revenue = 0;

  best.rows.forEach((row) => {
    const rowCountry = normalizeCountry(countryHeader ? row[countryHeader] : "");
    const rowStoreType = storeTypeHeader ? String(row[storeTypeHeader] ?? "").trim() : "";
    const rowOnlinePlatform = onlinePlatformHeader ? String(row[onlinePlatformHeader] ?? "").trim() : "";
    const isVietnamFranchiseRetail =
      kind === "sales" &&
      /越南|vietnam/i.test(rowCountry) &&
      /加盟/.test(rowStoreType);
    if (isVietnamFranchiseRetail) return;
    const qty = quantityHeader ? numeric(row[quantityHeader]) : 0;
    const rowAmount = valueHeader ? numeric(row[valueHeader]) : 0;
    const rowRevenue = revenueHeader ? numeric(row[revenueHeader]) : rowAmount;
    quantity += qty;
    amount += rowAmount;
    revenue += rowRevenue;
    add(sku, skuHeader ? String(row[skuHeader] ?? "").trim() : "", qty);
    add(countries, rowCountry, qty);
    add(monthly, dateHeader ? monthOf(row[dateHeader]) : "", qty);
    const month = dateHeader ? monthOf(row[dateHeader]) : "";
    const skuValue = skuHeader ? String(row[skuHeader] ?? "").trim() : "";
    const country = rowCountry;
    const name = nameHeader ? String(row[nameHeader] ?? "").trim() : "";
    const brand = brandHeader ? normalizeBrand(row[brandHeader]) : "";
    const category = categoryHeader ? String(row[categoryHeader] ?? "").trim() : "";
    const middleCategory = middleCategoryHeader ? String(row[middleCategoryHeader] ?? "").trim() : "";
    const smallCategory = smallCategoryHeader ? String(row[smallCategoryHeader] ?? "").trim() : "";
    const series = seriesHeader ? String(row[seriesHeader] ?? "").trim() : "";
    const rawChannel = channelHeader ? String(row[channelHeader] ?? "").trim() : "";
    const channel = kind === "wholesale"
      ? "批发"
      : kind === "sales"
        ? (rowOnlinePlatform || /电商|线上|shopee|lazada|tiktok|online|e-?commerce/i.test(rawChannel)
            ? "电商"
            : /自营/.test(rowStoreType)
              ? "自营"
              : "联营")
        : rawChannel;
    const analysisSku = skuValue || (isSalesKind(kind) ? `无SKU维度:${brand || "未识别品牌"}:${category || "未识别大类"}:${middleCategory || "未识别中类"}:${channel || "未识别渠道"}` : "");
    if (analysisSku) {
      const key = [month, country, analysisSku, name, brand, category, middleCategory, smallCategory, series, channel].join("¦");
      const current = detailMap.get(key);
      if (current) { current.quantity += qty; current.amount += rowAmount; current.revenue += rowRevenue; }
      else detailMap.set(key, { month, sku: analysisSku, name, brand, country, category, middleCategory, smallCategory, series, channel, sourceType: kind === "wholesale" ? "批发" : kind === "sales" ? "零售" : undefined, quantity: qty, amount: rowAmount, revenue: rowRevenue });
    }
  });
  const months = Object.keys(monthly).sort();
  const missing = [
    !quantityHeader && "数量",
    !skuHeader && !isSalesKind(kind) && "SKU",
    !dateHeader && "日期/月度",
  ].filter(Boolean);

  return {
    schemaVersion: schemaVersionFor(kind),
    kind,
    fileName,
    uploadedAt: new Date().toISOString(),
    rowCount: best.rows.filter((row) => Object.values(row).some((value) => String(value).trim())).length,
    sheetName: best.sheetName,
    headers: best.headers,
    status: missing.length ? "mapping" : "ready",
    message: missing.length ? `已保存，但未识别：${missing.join("、")}` : !skuHeader && isSalesKind(kind) ? `${kind === "wholesale" ? "批发" : "零售"}底稿无SKU，将按品牌、品类、国家和月份分析` : "字段已自动识别，可用于滚动分析",
    summary: {
      quantity,
      amount,
      revenue,
      skuCount: Object.keys(sku).length,
      monthCount: months.length,
      countryCount: Object.keys(countries).length,
      latestMonth: months.at(-1) ?? "",
      detectedDateField: dateHeader ?? "",
    },
    monthly,
    sku,
    countries,
    details: Array.from(detailMap.values()),
  };
}

function fmt(value: number, digits = 0) {
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: digits }).format(value);
}

function mergeSalesSources(retail?: Dataset, wholesale?: Dataset): Dataset | undefined {
  const sources = [retail, wholesale].filter(Boolean) as Dataset[];
  if (!sources.length) return undefined;
  const mergeMap = (field: "monthly" | "sku" | "countries") => {
    const result: Record<string, number> = {};
    sources.forEach((source) => Object.entries(source[field]).forEach(([key, value]) => add(result, key, value)));
    return result;
  };
  const monthly = mergeMap("monthly");
  const sku = mergeMap("sku");
  const countries = mergeMap("countries");
  return {
    schemaVersion: SALES_SCHEMA_VERSION,
    kind: "sales",
    fileName: sources.map((source) => source.fileName).join(" + "),
    uploadedAt: sources.map((source) => source.uploadedAt).sort().at(-1) ?? "",
    rowCount: sources.reduce((sum, source) => sum + source.rowCount, 0),
    sheetName: "零售 + 批发",
    headers: Array.from(new Set(sources.flatMap((source) => source.headers))),
    status: sources.every((source) => source.status === "ready") ? "ready" : "mapping",
    message: `已合并 ${retail ? "零售" : ""}${retail && wholesale ? " + " : ""}${wholesale ? "批发" : ""}销售`,
    summary: {
      quantity: sources.reduce((sum, source) => sum + source.summary.quantity, 0),
      amount: sources.reduce((sum, source) => sum + (source.summary.amount ?? 0), 0),
      revenue: sources.reduce((sum, source) => sum + (source.summary.revenue ?? source.summary.amount ?? 0), 0),
      skuCount: Object.keys(sku).length,
      monthCount: Object.keys(monthly).length,
      countryCount: Object.keys(countries).length,
      latestMonth: Object.keys(monthly).sort().at(-1) ?? "",
      detectedDateField: sources.map((source) => source.summary.detectedDateField).filter(Boolean).join(" / "),
    },
    monthly,
    sku,
    countries,
    details: sources.flatMap((source) => source.details ?? []),
  };
}

function mergeInventorySnapshots(existing: Dataset | undefined, incoming: Dataset): Dataset {
  const incomingMonths = new Set((incoming.details ?? []).map((row) => row.month).filter(Boolean));
  if (!incomingMonths.size) throw new Error("未识别库存日期，请确认底稿包含库存日期或月份字段");
  const details = [
    ...(existing?.details ?? []).filter((row) => !incomingMonths.has(row.month)),
    ...(incoming.details ?? []),
  ];
  const monthly: Record<string, number> = {};
  const sku: Record<string, number> = {};
  const countries: Record<string, number> = {};
  let quantity = 0;
  let amount = 0;
  details.forEach((row) => {
    quantity += row.quantity;
    amount += row.amount ?? 0;
    add(monthly, row.month, row.quantity);
    add(sku, row.sku, row.quantity);
    add(countries, row.country, row.quantity);
  });
  const months = Object.keys(monthly).sort();
  return {
    ...incoming,
    fileName: `${months.length}个月末库存快照`,
    uploadedAt: new Date().toISOString(),
    rowCount: details.length,
    headers: Array.from(new Set([...(existing?.headers ?? []), ...incoming.headers])),
    message: `已保存 ${months.length} 个月末库存；重复月份自动覆盖`,
    summary: {
      quantity,
      amount,
      skuCount: Object.keys(sku).length,
      monthCount: months.length,
      countryCount: Object.keys(countries).length,
      latestMonth: months.at(-1) ?? "",
      detectedDateField: incoming.summary.detectedDateField || existing?.summary.detectedDateField || "",
    },
    monthly,
    sku,
    countries,
    details,
  };
}

function nextMonths(start: string, count = 6) {
  const base = start ? new Date(`${start}-01T00:00:00`) : new Date();
  return Array.from({ length: count }, (_, index) => {
    const date = new Date(base.getFullYear(), base.getMonth() + index + 1, 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  });
}

function shiftMonth(month: string, offset: number) {
  if (!month) return "";
  const date = new Date(`${month}-01T00:00:00`);
  date.setMonth(date.getMonth() + offset);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

export default function DataHub({ view = "data" }: { view?: "data" | "analytics" | "sales" | "inventory" }) {
  const [datasets, setDatasets] = useState<Partial<Record<DataKind, Dataset>>>({});
  const [busy, setBusy] = useState<DataKind | null>(null);
  const [notice, setNotice] = useState("正在读取本机已保存的数据…");
  const [analysisMonth, setAnalysisMonth] = useState("");
  const [salesStartMonth, setSalesStartMonth] = useState("");
  const [inventoryAnalysisMonth, setInventoryAnalysisMonth] = useState("");
  const [countryFilter, setCountryFilter] = useState("全部");
  const [categoryFilter, setCategoryFilter] = useState("全部");
  const [brandFilter, setBrandFilter] = useState("全部");
  const [salesSourceFilter, setSalesSourceFilter] = useState("全部销售");
  const [middleCategoryFilter, setMiddleCategoryFilter] = useState("全部");
  const [smallCategoryFilter, setSmallCategoryFilter] = useState("全部");
  const [seriesFilter, setSeriesFilter] = useState("全部");
  const [skuSearch, setSkuSearch] = useState("");
  const [riskFilter, setRiskFilter] = useState("全部");
  const [metricMode, setMetricMode] = useState<"quantity" | "amount" | "revenue">("quantity");
  const [overviewBrand, setOverviewBrand] = useState("全部品牌");
  const inputs = useRef<Partial<Record<DataKind, HTMLInputElement | null>>>({});

  useEffect(() => {
    getAll()
      .then(async (rows) => {
        const validRows = rows.filter((row) => row.schemaVersion === schemaVersionFor(row.kind));
        const expiredRows = rows.filter((row) => row.schemaVersion !== schemaVersionFor(row.kind));
        if (expiredRows.length) await Promise.all(expiredRows.map((row) => deleteOne(row.kind)));
        setDatasets(Object.fromEntries(validRows.map((row) => [row.kind, row])));
        setNotice(expiredRows.length
          ? "旧版销售渠道口径已自动清除：请重新上传 R01 和批发销售；库存与 OIH 已保留"
          : validRows.length ? `已恢复 ${validRows.length} 类数据，无需再次上传` : "尚未上传真实数据，可从三个数据源开始");
      })
      .catch(() => setNotice("浏览器存储不可用，请检查隐私模式或网站存储权限"));
  }, []);

  const sales = useMemo(() => mergeSalesSources(datasets.sales, datasets.wholesale), [datasets.sales, datasets.wholesale]);
  const inventory = datasets.inventory;
  const oih = datasets.oih;
  const salesMonths = Object.entries(sales?.monthly ?? {}).sort(([a], [b]) => a.localeCompare(b));
  const recentSales = salesMonths.slice(-3);
  const monthlyVelocity = recentSales.length ? recentSales.reduce((sum, [, value]) => sum + value, 0) / recentSales.length : 0;
  const inventoryQty = inventory?.summary.quantity ?? 0;
  const oihQty = oih?.summary.quantity ?? 0;
  const activeSku = new Set([...Object.keys(sales?.sku ?? {}), ...Object.keys(inventory?.sku ?? {}), ...Object.keys(oih?.sku ?? {})]).size;
  const inventorySku = inventory?.summary.skuCount ?? 0;
  const depth = inventorySku ? inventoryQty / inventorySku : 0;
  const woi = monthlyVelocity ? (inventoryQty / monthlyVelocity) * 4.33 : 0;
  const detailReady = view === "sales"
    ? Boolean(sales?.details?.length)
    : view === "inventory"
      ? Boolean(inventory?.details?.length)
      : Boolean(sales?.details?.length && inventory?.details?.length);
  const availableMonths = Array.from(new Set((sales?.details ?? []).map((row) => row.month).filter(Boolean))).sort();
  const selectedMonth = analysisMonth || availableMonths.at(-1) || "";
  const salesRangeEnd = selectedMonth;
  const defaultSalesStart = availableMonths.filter((month) => month <= salesRangeEnd).slice(-12).at(0) ?? salesRangeEnd;
  const salesRangeStart = salesStartMonth && salesStartMonth <= salesRangeEnd ? salesStartMonth : defaultSalesStart;
  const salesRangeMonths = availableMonths.filter((month) => month >= salesRangeStart && month <= salesRangeEnd);
  const recentAnalysisMonths = selectedMonth ? availableMonths.filter((month) => month <= selectedMonth).slice(-3) : [];
  const latestInventoryMonth = Array.from(new Set((inventory?.details ?? []).map((row) => row.month).filter(Boolean))).sort().at(-1) ?? "";
  const inventorySnapshotMonths = Array.from(new Set((inventory?.details ?? []).map((row) => row.month).filter(Boolean))).sort().reverse();
  const inventoryViewMonth = view === "inventory" ? (inventoryAnalysisMonth || latestInventoryMonth) : latestInventoryMonth;
  const filterCountries = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.country).filter(Boolean))).sort();
  const filterCategories = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.category).filter(Boolean))).sort();
  const presentBrands = new Set(
    [...(sales?.details ?? []), ...(inventory?.details ?? [])]
      .map((row) => row.brand.trim())
      .filter(Boolean),
  );
  const filterBrands = CANONICAL_BRANDS.filter((brand) => presentBrands.has(brand));
  const salesDashboardRows = (sales?.details ?? []).filter((row) =>
    (salesSourceFilter === "全部销售" || row.sourceType === salesSourceFilter) &&
    (countryFilter === "全部" || row.country === countryFilter) &&
    (brandFilter === "全部" || row.brand === brandFilter) &&
    (categoryFilter === "全部" || row.category === categoryFilter) &&
    (middleCategoryFilter === "全部" || row.middleCategory === middleCategoryFilter) &&
    (smallCategoryFilter === "全部" || row.smallCategory === smallCategoryFilter) &&
    (seriesFilter === "全部" || row.series === seriesFilter)
  );
  const salesCurrentRows = salesDashboardRows.filter((row) =>
    view === "sales" ? salesRangeMonths.includes(row.month) : row.month === selectedMonth
  );
  const salesRevenue = salesCurrentRows.reduce((sum, row) => sum + (row.revenue ?? row.amount ?? 0), 0);
  const salesTicketAmount = salesCurrentRows.reduce((sum, row) => sum + (row.amount ?? 0), 0);
  const salesQuantity = salesCurrentRows.reduce((sum, row) => sum + row.quantity, 0);
  const periodLength = Math.max(salesRangeMonths.length, 1);
  const previousMonth = view === "sales" ? shiftMonth(salesRangeStart, -1) : availableMonths.filter((month) => month < selectedMonth).at(-1) ?? "";
  const previousPeriodStart = view === "sales" ? shiftMonth(salesRangeStart, -periodLength) : previousMonth;
  const previousPeriodEnd = previousMonth;
  const previousRevenue = salesDashboardRows
    .filter((row) => row.month >= previousPeriodStart && row.month <= previousPeriodEnd)
    .reduce((sum, row) => sum + (row.revenue ?? row.amount ?? 0), 0);
  const revenueGrowth = previousRevenue ? salesRevenue / previousRevenue - 1 : 0;
  const priorYearStart = shiftMonth(view === "sales" ? salesRangeStart : selectedMonth, -12);
  const priorYearMonth = shiftMonth(view === "sales" ? salesRangeEnd : selectedMonth, -12);
  const priorYearRevenue = salesDashboardRows
    .filter((row) => row.month >= priorYearStart && row.month <= priorYearMonth)
    .reduce((sum, row) => sum + (row.revenue ?? row.amount ?? 0), 0);
  const revenueYoY = priorYearRevenue ? salesRevenue / priorYearRevenue - 1 : 0;
  const salesMonthlyRevenue = (view === "sales" ? salesRangeMonths : availableMonths.slice(-12)).map((month) => ({
    name: month,
    value: salesDashboardRows.filter((row) => row.month === month).reduce((sum, row) => sum + (row.revenue ?? row.amount ?? 0), 0) / 10000,
  }));
  const salesGroup = (field: "country" | "channel" | "brand" | "category") => {
    const grouped = new Map<string, number>();
    salesCurrentRows.forEach((row) => {
      const name = row[field] || "未识别";
      grouped.set(name, (grouped.get(name) ?? 0) + (row.revenue ?? row.amount ?? 0));
    });
    return Array.from(grouped, ([name, value]) => ({ name, value: value / 10000 })).sort((a, b) => b.value - a.value).slice(0, 12);
  };
  const countryRevenue = salesGroup("country");
  const channelRevenue = salesGroup("channel");
  const brandRevenue = salesGroup("brand");
  const categoryRevenue = salesGroup("category");
  const overviewMatches = (row: DetailRecord) => overviewBrand === "全部品牌" || row.brand === overviewBrand;
  const overviewInventoryRows = (inventory?.details ?? []).filter((row) => overviewMatches(row) && (!latestInventoryMonth || !row.month || row.month === latestInventoryMonth));
  const overviewSalesRows = (sales?.details ?? []).filter(overviewMatches);
  const overviewOihRows = (oih?.details ?? []).filter(overviewMatches);
  const overviewSalesMonths = Array.from(new Set(overviewSalesRows.map((row) => row.month).filter(Boolean))).sort();
  const overviewLatestSalesMonth = overviewSalesMonths.at(-1) ?? "";
  const overviewRecentMonths = overviewSalesMonths.slice(-3);
  const overviewInventoryAmount = overviewInventoryRows.reduce((sum, row) => sum + (row.amount ?? 0), 0);
  const overviewLatestSalesAmount = overviewSalesRows.filter((row) => row.month === overviewLatestSalesMonth).reduce((sum, row) => sum + (row.amount ?? 0), 0);
  const overviewVelocityAmount = overviewRecentMonths.length
    ? overviewSalesRows.filter((row) => overviewRecentMonths.includes(row.month)).reduce((sum, row) => sum + (row.amount ?? 0), 0) / overviewRecentMonths.length
    : 0;
  const overviewOihAmount = overviewOihRows.reduce((sum, row) => sum + (row.amount ?? 0), 0);
  const overviewAmountRatio = overviewVelocityAmount ? overviewInventoryAmount / overviewVelocityAmount : 0;
  const overviewSalesTrend = overviewSalesMonths.slice(-12).map((month) => ({
    month,
    value: overviewSalesRows.filter((row) => row.month === month).reduce((sum, row) => sum + (row.amount ?? 0), 0) / 10000,
  }));
  const filterMiddleCategories = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].filter((row) => categoryFilter === "全部" || row.category === categoryFilter).map((row) => row.middleCategory).filter(Boolean))).sort();
  const filterSmallCategories = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].filter((row) => (categoryFilter === "全部" || row.category === categoryFilter) && (middleCategoryFilter === "全部" || row.middleCategory === middleCategoryFilter)).map((row) => row.smallCategory).filter(Boolean))).sort();
  const filterSeries = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.series).filter(Boolean))).sort();

  const skuAnalysis = useMemo(() => {
    type SkuLine = { key: string; sku: string; name: string; brand: string; country: string; category: string; middleCategory: string; smallCategory: string; series: string; channel: string; stock: number; stockAmount: number; sales: number; salesAmount: number; sales3m: number; sales3mAmount: number; oih: number; oihAmount: number; ratio: number; amountRatio: number; weeks: number; status: string };
    const lines = new Map<string, SkuLine>();
    const ensure = (row: DetailRecord) => {
      const key = `${row.country}¦${row.sku}`;
      if (!lines.has(key)) lines.set(key, { key, sku: row.sku, name: row.name, brand: row.brand, country: row.country, category: row.category, middleCategory: row.middleCategory, smallCategory: row.smallCategory, series: row.series, channel: row.channel, stock: 0, stockAmount: 0, sales: 0, salesAmount: 0, sales3m: 0, sales3mAmount: 0, oih: 0, oihAmount: 0, ratio: 0, amountRatio: 0, weeks: 0, status: "" });
      const line = lines.get(key)!;
      if (!line.name && row.name) line.name = row.name;
      if (!line.category && row.category) line.category = row.category;
      if (!line.brand && row.brand) line.brand = row.brand;
      if (!line.middleCategory && row.middleCategory) line.middleCategory = row.middleCategory;
      if (!line.smallCategory && row.smallCategory) line.smallCategory = row.smallCategory;
      if (!line.series && row.series) line.series = row.series;
      if (!line.channel && row.channel) line.channel = row.channel;
      return line;
    };
    (inventory?.details ?? []).filter((row) => !inventoryViewMonth || !row.month || row.month === inventoryViewMonth).forEach((row) => { const line = ensure(row); line.stock += row.quantity; line.stockAmount += row.amount ?? 0; });
    (sales?.details ?? []).filter((row) =>
      recentAnalysisMonths.includes(row.month) &&
      (salesSourceFilter === "全部销售" || row.sourceType === salesSourceFilter)
    ).forEach((row) => {
      const line = ensure(row);
      line.sales3m += row.quantity / Math.max(recentAnalysisMonths.length, 1);
      line.sales3mAmount += (row.amount ?? 0) / Math.max(recentAnalysisMonths.length, 1);
      if (row.month === selectedMonth) { line.sales += row.quantity; line.salesAmount += row.amount ?? 0; }
    });
    (oih?.details ?? []).forEach((row) => { const line = ensure(row); line.oih += row.quantity; line.oihAmount += row.amount ?? 0; });
    return Array.from(lines.values()).map((line) => {
      line.ratio = line.sales3m > 0 ? line.stock / line.sales3m : line.stock > 0 ? 999 : 0;
      line.amountRatio = line.sales3mAmount > 0 ? line.stockAmount / line.sales3mAmount : line.stockAmount > 0 ? 999 : 0;
      line.weeks = line.sales3m > 0 ? line.ratio * 4.33 : line.stock > 0 ? 999 : 0;
      line.status = line.stock > 0 && line.sales3m <= 0 ? "无动销" : line.ratio > 6 ? "高库存" : line.ratio < 1 && line.sales3m > 0 ? "缺货风险" : "健康";
      return line;
    }).filter((line) =>
      (countryFilter === "全部" || line.country === countryFilter) &&
      (brandFilter === "全部" || line.brand === brandFilter) &&
      (categoryFilter === "全部" || line.category === categoryFilter) &&
      (middleCategoryFilter === "全部" || line.middleCategory === middleCategoryFilter) &&
      (smallCategoryFilter === "全部" || line.smallCategory === smallCategoryFilter) &&
      (seriesFilter === "全部" || line.series === seriesFilter) &&
      (riskFilter === "全部" || line.status === riskFilter) &&
      (!skuSearch || `${line.sku} ${line.name}`.toLowerCase().includes(skuSearch.toLowerCase())),
    ).sort((a, b) => b.stock - a.stock);
  }, [sales?.details, inventory?.details, oih?.details, inventoryViewMonth, selectedMonth, recentAnalysisMonths.join("|"), salesSourceFilter, countryFilter, brandFilter, categoryFilter, middleCategoryFilter, smallCategoryFilter, seriesFilter, riskFilter, skuSearch]);

  const analyticStock = skuAnalysis.reduce((sum, row) => sum + row.stock, 0);
  const analyticSales = skuAnalysis.reduce((sum, row) => sum + row.sales, 0);
  const analyticVelocity = skuAnalysis.reduce((sum, row) => sum + row.sales3m, 0);
  const analyticStockAmount = skuAnalysis.reduce((sum, row) => sum + row.stockAmount, 0);
  const analyticSalesAmount = skuAnalysis.reduce((sum, row) => sum + row.salesAmount, 0);
  const analyticVelocityAmount = skuAnalysis.reduce((sum, row) => sum + row.sales3mAmount, 0);
  const analyticRatio = analyticVelocity ? analyticStock / analyticVelocity : 0;
  const analyticAmountRatio = analyticVelocityAmount ? analyticStockAmount / analyticVelocityAmount : 0;
  const stockedSku = skuAnalysis.filter((row) => row.stock > 0).length;
  const movingSku = skuAnalysis.filter((row) => row.sales3m > 0).length;
  const zeroMovingStock = skuAnalysis.filter((row) => row.status === "无动销").reduce((sum, row) => sum + row.stock, 0);
  const healthyRate = stockedSku ? skuAnalysis.filter((row) => row.status === "健康").length / stockedSku : 0;
  const dimensionSummary = (field: "brand" | "category" | "middleCategory" | "smallCategory" | "series" | "country") => {
    const grouped = new Map<string, { name: string; stock: number; sales: number; stockAmount: number; salesAmount: number }>();
    skuAnalysis.forEach((row) => {
      const name = row[field] || "未识别";
      const current = grouped.get(name) ?? { name, stock: 0, sales: 0, stockAmount: 0, salesAmount: 0 };
      current.stock += row.stock;
      current.sales += row.sales3m;
      current.stockAmount += row.stockAmount;
      current.salesAmount += row.sales3mAmount;
      grouped.set(name, current);
    });
    return Array.from(grouped.values()).sort((a, b) => b.stock - a.stock).slice(0, 10);
  };
  const brandMix = dimensionSummary("brand");
  const categoryMix = dimensionSummary("category");
  const countryMix = dimensionSummary("country");
  const seriesMix = dimensionSummary("series");

  const rolling = useMemo(() => {
    const months = nextMonths(sales?.summary.latestMonth || inventory?.summary.latestMonth || "", 6);
    let opening = overviewInventoryAmount / 10000;
    const forecast = overviewVelocityAmount / 10000;
    return months.map((month, index) => {
      const arrivals = index === 1 ? overviewOihAmount / 10000 : 0;
      const closing = opening + arrivals - forecast;
      const row = { month, opening, arrivals, sales: forecast, closing };
      opening = closing;
      return row;
    });
  }, [overviewInventoryAmount, overviewVelocityAmount, overviewOihAmount, sales?.summary.latestMonth, inventory?.summary.latestMonth]);

  async function upload(kind: DataKind, event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setBusy(kind);
    setNotice(`正在解析 ${file.name}…`);
    try {
      const parsed = parseWorkbook(await file.arrayBuffer(), kind, file.name);
      const saved = kind === "inventory" ? mergeInventorySnapshots(datasets.inventory, parsed) : parsed;
      await putOne(saved);
      setDatasets((current) => ({ ...current, [kind]: saved }));
      setNotice(kind === "inventory"
        ? `${file.name} 已按月份追加；相同月份已自动覆盖`
        : `${file.name} 已保存；下次打开会自动恢复`);
    } catch (error) {
      setNotice(error instanceof Error ? `读取失败：${error.message}` : "读取失败，请检查文件格式");
    } finally {
      setBusy(null);
      event.target.value = "";
    }
  }

  async function remove(kind: DataKind) {
    await deleteOne(kind);
    setDatasets((current) => {
      const next = { ...current };
      delete next[kind];
      return next;
    });
    setNotice("已删除该类本机数据，可随时重新上传");
  }

  async function removeInventoryMonth(month: string) {
    const current = datasets.inventory;
    if (!current) return;
    const remaining = (current.details ?? []).filter((row) => row.month !== month);
    if (!remaining.length) {
      await remove("inventory");
      return;
    }
    const placeholder: Dataset = {
      ...current,
      details: remaining,
      monthly: {},
      sku: {},
      countries: {},
      summary: { ...current.summary, quantity: 0, amount: 0, skuCount: 0, monthCount: 0, countryCount: 0, latestMonth: "" },
    };
    const rebuilt = mergeInventorySnapshots(undefined, placeholder);
    await putOne(rebuilt);
    setDatasets((datasetsNow) => ({ ...datasetsNow, inventory: rebuilt }));
    setNotice(`已删除 ${month} 月末库存快照`);
  }

  function exportBackup() {
    const blob = new Blob([JSON.stringify({ version: 1, exportedAt: new Date().toISOString(), datasets }, null, 2)], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Wilson数据备份_${new Date().toISOString().slice(0, 10)}.json`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  async function importBackup(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as { datasets?: Partial<Record<DataKind, Dataset>> };
      const rows = Object.values(parsed.datasets ?? {}).filter(Boolean) as Dataset[];
      await Promise.all(rows.map(putOne));
      setDatasets(Object.fromEntries(rows.map((row) => [row.kind, row])));
      setNotice(`备份恢复成功，共 ${rows.length} 类数据`);
    } catch {
      setNotice("备份恢复失败：请选择由本看板导出的 JSON 文件");
    }
    event.target.value = "";
  }

  const conclusion = !detailReady
    ? "请按新口径重新上传销售与月末库存，系统将按品牌和RMB吊牌金额重算月度库销比。"
    : overviewAmountRatio > 6
      ? `${overviewBrand}库存吊牌金额约为近3月月均销售额的 ${overviewAmountRatio.toFixed(1)} 倍，偏高；建议先消化库存，并将新订单拆分到货。`
      : rolling.some((row) => row.closing < 0)
        ? `${overviewBrand}按当前销售额滚动预测将出现库存不足；建议结合到货月份补单。`
        : `${overviewBrand}金额库销比为 ${overviewAmountRatio.toFixed(1)} 个月，滚动期内库存可控；可继续按品类和SKU下钻审核。`;

  if (view === "analytics" || view === "sales" || view === "inventory") {
    const viewTitle = view === "sales" ? "销售经营看板" : view === "inventory" ? "库存经营看板" : "进销存经营总览";
    const viewEyebrow = view === "sales" ? "SALES PERFORMANCE" : view === "inventory" ? "INVENTORY MANAGEMENT" : "SALES & INVENTORY OVERVIEW";
    const viewDescription = view === "sales"
      ? "按月度、销售类型、品牌、国家和品类查看销售趋势、结构与动销表现。"
      : view === "inventory"
        ? "按品牌、国家、品类、系列和SKU查看库存规模、结构、深度与积压风险。"
        : "将零售、批发、库存与OIH放在一起，查看月度库销关系及经营风险。";
    return (
      <section className="data-hub analytics-view">
        <div className="page-heading">
          <div><p className="eyebrow blue">{viewEyebrow}</p><h2>{viewTitle}</h2><p>{viewDescription} 上传数据仍保存在当前浏览器。</p></div>
          <div className="local-badge"><i /> 同一设备自动记住</div>
        </div>
        <div className="analytics-filterbar">
          {view === "sales" ? <>
            <label>开始月份<select value={salesRangeStart} onChange={(event) => setSalesStartMonth(event.target.value)}>{availableMonths.filter((month) => month <= salesRangeEnd).map((month) => <option key={month}>{month}</option>)}</select></label>
            <label>结束月份<select value={salesRangeEnd} onChange={(event) => setAnalysisMonth(event.target.value)}>{availableMonths.filter((month) => month >= salesRangeStart).map((month) => <option key={month}>{month}</option>)}</select></label>
          </> : view === "inventory"
            ? <label>库存快照月份<select value={inventoryViewMonth} onChange={(event) => setInventoryAnalysisMonth(event.target.value)}>{inventorySnapshotMonths.map((month) => <option key={month}>{month}</option>)}</select></label>
            : <label>分析月份<select value={selectedMonth} onChange={(event) => setAnalysisMonth(event.target.value)}><option value="">最新月份</option>{availableMonths.map((month) => <option key={month}>{month}</option>)}</select></label>}
          <label>销售类型<select value={salesSourceFilter} onChange={(event) => setSalesSourceFilter(event.target.value)}><option>全部销售</option><option>零售</option><option>批发</option></select></label>
          <label>品牌<select value={brandFilter} onChange={(event) => setBrandFilter(event.target.value)}><option>全部</option>{filterBrands.map((brand) => <option key={brand}>{brand}</option>)}</select></label>
          <label>国家<select value={countryFilter} onChange={(event) => setCountryFilter(event.target.value)}><option>全部</option>{filterCountries.map((country) => <option key={country}>{country}</option>)}</select></label>
          <label>大类<select value={categoryFilter} onChange={(event) => { setCategoryFilter(event.target.value); setMiddleCategoryFilter("全部"); setSmallCategoryFilter("全部"); }}><option>全部</option>{filterCategories.map((category) => <option key={category}>{category}</option>)}</select></label>
          <label>中类<select value={middleCategoryFilter} onChange={(event) => { setMiddleCategoryFilter(event.target.value); setSmallCategoryFilter("全部"); }}><option>全部</option>{filterMiddleCategories.map((category) => <option key={category}>{category}</option>)}</select></label>
          <label>小类<select value={smallCategoryFilter} onChange={(event) => setSmallCategoryFilter(event.target.value)}><option>全部</option>{filterSmallCategories.map((category) => <option key={category}>{category}</option>)}</select></label>
          <label>系列<select value={seriesFilter} onChange={(event) => setSeriesFilter(event.target.value)}><option>全部</option>{filterSeries.map((series) => <option key={series}>{series}</option>)}</select></label>
          {view !== "sales" && <label>库存状态<select value={riskFilter} onChange={(event) => setRiskFilter(event.target.value)}>{["全部", "健康", "高库存", "无动销", "缺货风险"].map((status) => <option key={status}>{status}</option>)}</select></label>}
          <label className="search-filter">SKU / 商品<input value={skuSearch} onChange={(event) => setSkuSearch(event.target.value)} placeholder="输入货号或品名" /></label>
          <div className="metric-switch"><span>图表口径</span><button className={metricMode === "quantity" ? "active" : ""} onClick={() => setMetricMode("quantity")}>数量</button><button className={metricMode === "amount" ? "active" : ""} onClick={() => setMetricMode("amount")}>吊牌金额</button>{view === "sales" && <button className={metricMode === "revenue" ? "active" : ""} onClick={() => setMetricMode("revenue")}>流水</button>}</div>
        </div>
        {view === "sales" && <div className="detail-upgrade-note"><strong>销售经营口径</strong><span>R01 零售数据已自动剔除“越南 + 加盟店”，该部分属于批发客户终端销售，避免与批发流水重复计算。</span></div>}
        {!detailReady && <div className="detail-upgrade-note"><strong>需要重新上传销售与库存文件</strong><span>旧版只保存总量汇总；新版上传后会保留SKU、国家、品类、渠道和月份维度，才能生成细颗粒度分析。</span></div>}
        <div className="summary-grid analytics-kpis">
          {view !== "inventory" && <>
          <article><span>{view === "sales" ? "区间销售流水" : "当月销售流水"}</span><strong>{detailReady ? fmt(salesRevenue / 10000, 1) : "—"} <em>万元</em></strong><small>{view === "sales" ? `${salesRangeStart} 至 ${salesRangeEnd}` : selectedMonth || "等待销售月份"} · RMB实际销售额</small></article>
          <article><span>{view === "sales" ? "较上一区间" : "流水环比"}</span><strong className={revenueGrowth < 0 ? "red" : ""}>{detailReady && previousRevenue ? `${revenueGrowth >= 0 ? "+" : ""}${(revenueGrowth * 100).toFixed(1)}%` : "—"}</strong><small>{view === "sales" ? `对比 ${previousPeriodStart} 至 ${previousPeriodEnd}` : `对比 ${previousMonth || "上月"}`}</small></article>
          <article><span>流水同比</span><strong className={revenueYoY < 0 ? "red" : ""}>{detailReady && priorYearRevenue ? `${revenueYoY >= 0 ? "+" : ""}${(revenueYoY * 100).toFixed(1)}%` : "—"}</strong><small>{priorYearRevenue ? `对比 ${priorYearStart}${view === "sales" && priorYearStart !== priorYearMonth ? ` 至 ${priorYearMonth}` : ""}` : "无去年同期数据"}</small></article>
          <article><span>{view === "sales" ? "区间销售吊牌金额" : "当月销售吊牌金额"}</span><strong>{detailReady ? fmt(salesTicketAmount / 10000, 1) : "—"} <em>万元</em></strong><small>RMB吊牌口径</small></article>
          <article><span>{view === "sales" ? "区间销量" : "当月销量"}</span><strong>{detailReady ? fmt(salesQuantity) : "—"} <em>件</em></strong><small>{salesSourceFilter}</small></article>
          </>}
          {view !== "sales" && <>
          <article><span>期末库存</span><strong>{detailReady ? fmt(analyticStock) : "—"} <em>件</em></strong><small>{inventoryViewMonth || "等待库存月份"}</small></article>
          <article><span>库销比</span><strong className={analyticRatio > 6 ? "red" : analyticRatio > 4 ? "amber" : ""}>{detailReady ? analyticRatio.toFixed(1) : "—"} <em>月</em></strong><small>库存 ÷ 近3月月均销量</small></article>
          <article><span>动销率</span><strong>{detailReady && stockedSku ? `${((movingSku / stockedSku) * 100).toFixed(1)}%` : "—"}</strong><small>有销量SKU ÷ 有库存SKU</small></article>
          <article><span>库存宽度</span><strong>{detailReady ? fmt(stockedSku) : "—"} <em>SKU</em></strong><small>当前有库存的SKU数量</small></article>
          <article><span>平均深度</span><strong>{detailReady && stockedSku ? fmt(analyticStock / stockedSku, 1) : "—"} <em>件/SKU</em></strong><small>库存数量 ÷ 有库存SKU</small></article>
          <article><span>无动销库存</span><strong className="red">{detailReady ? fmt(zeroMovingStock) : "—"} <em>件</em></strong><small>近3月无销售但仍有库存</small></article>
          <article><span>健康SKU占比</span><strong>{detailReady ? `${(healthyRate * 100).toFixed(1)}%` : "—"}</strong><small>库销比1–6个月且持续动销</small></article>
          <article><span>库存吊牌金额</span><strong>{detailReady ? fmt(analyticStockAmount / 10000, 1) : "—"} <em>万元</em></strong><small>RMB吊牌口径</small></article>
          <article><span>金额库销比</span><strong className={analyticAmountRatio > 6 ? "red" : analyticAmountRatio > 4 ? "amber" : ""}>{detailReady ? analyticAmountRatio.toFixed(1) : "—"} <em>月</em></strong><small>库存吊牌额 ÷ 近3月月均销售吊牌额</small></article>
          </>}
        </div>
        {view === "sales" && <>
          <section className="panel sales-trend-panel">
            <div className="panel-title"><div><span className="step">01</span><h3>月度流水变化</h3></div><span className="unit">RMB流水 · 万元</span></div>
            <div className="mini-trend sales-revenue-trend">
              {salesMonthlyRevenue.map((row) => {
                const max = Math.max(...salesMonthlyRevenue.map((item) => item.value), 1);
                return <div key={row.name}><span>{row.name.slice(2)}</span><i style={{ height: `${Math.max(5, row.value / max * 100)}%` }} /><b>{fmt(row.value, 1)}</b></div>;
              })}
            </div>
          </section>
          <div className="dimension-grid sales-dimensions">
            {[
              ["国家流水规模", countryRevenue],
              ["渠道流水规模", channelRevenue],
              ["品牌流水规模", brandRevenue],
              ["大类流水规模", categoryRevenue],
            ].map(([title, rows]) => {
              const data = rows as Array<{ name: string; value: number }>;
              const max = Math.max(...data.map((row) => row.value), 1);
              return <section className="panel dimension-panel" key={title as string}>
                <div className="panel-title"><div><h3>{title as string}</h3></div><span className="unit">{selectedMonth} · 万元</span></div>
                <div className="sales-scale-list">{data.length ? data.map((row) => <div key={row.name}><strong>{row.name}</strong><i><b style={{ width: `${row.value / max * 100}%` }} /></i><span>{fmt(row.value, 1)}</span></div>) : <p>等待流水数据</p>}</div>
              </section>;
            })}
          </div>
        </>}
        {view !== "sales" && <>
        <div className="analytics-panels">
          <section className="panel status-distribution">
            <div className="panel-title"><div><span className="step">01</span><h3>库存状态分布</h3></div><span className="unit">SKU数</span></div>
            <div className="status-bars">
              {["健康", "高库存", "无动销", "缺货风险"].map((status) => {
                const count = skuAnalysis.filter((row) => row.status === status).length;
                return <div key={status}><span>{status}</span><i><b className={`risk-fill ${status}`} style={{ width: `${skuAnalysis.length ? Math.max(3, count / skuAnalysis.length * 100) : 0}%` }} /></i><strong>{count}</strong></div>;
              })}
            </div>
          </section>
          <section className="panel top-stock">
            <div className="panel-title"><div><span className="step">02</span><h3>库存占用 TOP 8</h3></div><span className="unit">件</span></div>
            <div className="top-stock-list">{skuAnalysis.slice(0, 8).map((row) => <div key={row.key}><span><b>{row.sku}</b><small>{row.country} · {row.name || "未识别品名"}</small></span><i><b style={{ width: `${analyticStock ? row.stock / Math.max(...skuAnalysis.map((item) => item.stock), 1) * 100 : 0}%` }} /></i><strong>{fmt(row.stock)}</strong></div>)}</div>
          </section>
        </div>
        <div className="dimension-grid">
          {[
            ["品牌库存与动销", brandMix],
            ["大类库存结构", categoryMix],
            ["国家库存分布", countryMix],
            ["系列库存结构", seriesMix],
          ].map(([title, rows]) => {
            const data = rows as Array<{ name: string; stock: number; sales: number; stockAmount: number; salesAmount: number }>;
            const maxStock = Math.max(...data.map((row) => metricMode === "amount" ? row.stockAmount : row.stock), 1);
            return <section className="panel dimension-panel" key={title as string}><div className="panel-title"><div><h3>{title as string}</h3></div><span className="legend"><span className="line-key actual" />库存 <span className="line-key target" />月均销售</span></div>
              <div className="dimension-list">{data.length ? data.map((row) => { const stockValue = metricMode === "amount" ? row.stockAmount / 10000 : row.stock; const salesValue = metricMode === "amount" ? row.salesAmount / 10000 : row.sales; const scaleStock = metricMode === "amount" ? row.stockAmount : row.stock; const scaleSales = metricMode === "amount" ? row.salesAmount : row.sales; return <div key={row.name}><strong>{row.name}</strong><div className="dual-bar"><i style={{ width: `${scaleStock / maxStock * 100}%` }} /><b style={{ width: `${Math.min(100, scaleSales / maxStock * 100)}%` }} /></div><span>{fmt(stockValue, metricMode === "amount" ? 1 : 0)} / {fmt(salesValue, 1)}{metricMode === "amount" ? "万" : ""}</span></div>; }) : <p>等待维度数据</p>}</div>
            </section>;
          })}
        </div>
        <section className="panel sku-diagnostic">
          <div className="panel-title"><div><span className="step">03</span><h3>SKU经营诊断明细</h3></div><span className="unit">{fmt(skuAnalysis.length)} 行 · 点击上方筛选</span></div>
          <div className="table-wrap"><table><thead><tr><th>SKU / 商品</th><th>品牌</th><th>国家</th><th>大/中/小类</th><th>系列 / 渠道</th><th>期末库存</th><th>库存吊牌额<br/>万元</th><th>当月销售</th><th>销售吊牌额<br/>万元</th><th>近3月月均</th><th>数量库销比</th><th>金额库销比</th><th>OIH</th><th>诊断</th></tr></thead>
          <tbody>{skuAnalysis.slice(0, 500).map((row) => <tr key={row.key}><td><strong>{row.sku}</strong><small>{row.name || "未识别品名"}</small></td><td>{row.brand || "未识别"}</td><td>{row.country || "未识别"}</td><td><strong>{row.category || "—"} / {row.middleCategory || "—"}</strong><small>{row.smallCategory || "未识别小类"}</small></td><td><strong>{row.series || "未识别系列"}</strong><small>{row.channel || "未识别渠道"}</small></td><td><strong>{fmt(row.stock)}</strong></td><td>{fmt(row.stockAmount / 10000, 1)}</td><td>{fmt(row.sales)}</td><td>{fmt(row.salesAmount / 10000, 1)}</td><td>{fmt(row.sales3m, 1)}</td><td>{row.ratio >= 999 ? "无销量" : `${row.ratio.toFixed(1)}月`}</td><td>{row.amountRatio >= 999 ? "无金额" : `${row.amountRatio.toFixed(1)}月`}</td><td>{fmt(row.oih)}</td><td><span className={`diagnostic-tag ${row.status}`}>{row.status}</span></td></tr>)}</tbody></table></div>
          {skuAnalysis.length > 500 && <p className="table-limit">当前显示库存最高的500行，请使用筛选缩小范围。</p>}
        </section>
        </>}
        <p className="privacy-note">计算口径：库存取最新可识别库存月份；当月销售取所选月份；数量与金额库销比均使用截至所选月份的最近3个月月均销售。金额统一采用RMB吊牌价并换算为万元（源表金额 ÷ 10,000）。</p>
      </section>
    );
  }

  return (
    <section className="data-hub">
      <div className="page-heading">
        <div>
          <p className="eyebrow blue">LOCAL DATA WORKSPACE</p>
          <h2>滚动进销存数据中心</h2>
          <p>上传一次后自动保存在当前浏览器。原始 Excel 不上传服务器，仅保存分析所需的汇总结果。</p>
        </div>
        <div className="local-badge"><i /> 同一设备自动记住</div>
      </div>

      <div className="data-toolbar">
        <span>{notice}</span>
        <div>
          <label className="overview-brand">品牌
            <select value={overviewBrand} onChange={(event) => setOverviewBrand(event.target.value)}>
              <option>全部品牌</option>
              {filterBrands.map((brand) => <option key={brand}>{brand}</option>)}
            </select>
          </label>
          <button className="secondary-button" onClick={exportBackup} disabled={!Object.keys(datasets).length}>导出数据备份</button>
          <label className="secondary-button file-button">恢复数据备份<input type="file" accept=".json" onChange={importBackup} /></label>
        </div>
      </div>

      <div className="upload-grid">
        {kinds.map((source) => {
          const data = datasets[source.kind];
          return (
            <article className={`upload-card ${source.accent}`} key={source.kind}>
              <div className="upload-title">
                <div><span>{source.title}</span><small>{source.subtitle}</small></div>
                <b className={data?.status === "ready" ? "ready" : data ? "mapping" : ""}>{data ? (data.status === "ready" ? "已就绪" : "待确认") : "未上传"}</b>
              </div>
              {data ? (
                <>
                  <strong className="file-name">{data.fileName}</strong>
                  <p>{fmt(data.rowCount)} 行 · {data.sheetName} · {new Date(data.uploadedAt).toLocaleString("zh-CN")}</p>
                  <div className="upload-stats"><span>数量 <b>{fmt(data.summary.quantity)}</b></span><span>{isSalesKind(source.kind) ? "流水" : "吊牌额"} <b>{fmt(((isSalesKind(source.kind) ? data.summary.revenue : data.summary.amount) ?? 0) / 10000, 1)}万</b></span><span>月份 <b>{fmt(data.summary.monthCount)}</b></span></div>
                  <small className="detected-field">月份字段：{data.summary.detectedDateField || "未识别"}</small>
                  {source.kind === "inventory" && <div className="snapshot-list">
                    <strong>已保存月末快照</strong>
                    <div>{inventorySnapshotMonths.map((month) => <span key={month}>{month}<button onClick={() => removeInventoryMonth(month)} aria-label={`删除 ${month} 库存`}>×</button></span>)}</div>
                  </div>}
                  <small className={data.status === "ready" ? "parse-ok" : "parse-warning"}>{data.message}</small>
                </>
              ) : <p className="empty-copy">支持 .xlsx、.xls、.csv；会自动寻找前25行中的表头并识别数量、SKU、国家和月份。</p>}
              <div className="upload-actions">
                <input ref={(node) => { inputs.current[source.kind] = node; }} hidden type="file" accept=".xlsx,.xls,.csv" onChange={(event) => upload(source.kind, event)} />
                <button className="primary small" disabled={busy !== null} onClick={() => inputs.current[source.kind]?.click()}>{busy === source.kind ? "正在解析…" : source.kind === "inventory" && data ? "追加月末快照" : data ? "覆盖上传" : "选择文件"}</button>
                {data && <button className="danger-link" onClick={() => remove(source.kind)}>删除</button>}
              </div>
            </article>
          );
        })}
      </div>

      <div className="summary-grid data-kpis">
        <article><span>库存吊牌金额</span><strong>{detailReady ? fmt(overviewInventoryAmount / 10000, 1) : "—"} <em>万元</em></strong><small>{overviewBrand} · RMB吊牌 · {latestInventoryMonth || "待上传"}</small></article>
        <article><span>当月销售吊牌金额</span><strong>{detailReady ? fmt(overviewLatestSalesAmount / 10000, 1) : "—"} <em>万元</em></strong><small>{overviewBrand} · {overviewLatestSalesMonth || "待上传"}</small></article>
        <article><span>近3月月均销售额</span><strong>{detailReady ? fmt(overviewVelocityAmount / 10000, 1) : "—"} <em>万元/月</em></strong><small>RMB吊牌口径</small></article>
        <article><span>金额库销比</span><strong className={overviewAmountRatio > 6 ? "red" : overviewAmountRatio > 4 ? "amber" : ""}>{detailReady ? overviewAmountRatio.toFixed(1) : "—"} <em>月</em></strong><small>库存吊牌额 ÷ 近3月月均销售吊牌额</small></article>
      </div>

      <section className="data-conclusion">
        <div><span>当前系统结论</span><h3>{conclusion}</h3></div>
        <strong>{detailReady ? (overviewAmountRatio > 6 ? "建议控单" : "进入明细审核") : "需重新上传"}</strong>
      </section>

      <div className="data-analysis-grid">
        <section className="panel">
          <div className="panel-title"><div><span className="step">01</span><h3>近月销售趋势</h3></div><span className="unit">RMB吊牌 · 万元</span></div>
          {overviewSalesTrend.length ? (
            <div className="mini-trend">
              {overviewSalesTrend.map(({ month, value }) => {
                const max = Math.max(...overviewSalesTrend.map((item) => item.value), 1);
                return <div key={month}><span>{month.slice(2)}</span><i style={{ height: `${Math.max(5, (value / max) * 100)}%` }} /><b>{fmt(value)}</b></div>;
              })}
            </div>
          ) : <div className="empty-panel">上传销售文件后显示最近12个月趋势</div>}
        </section>

        <section className="panel">
          <div className="panel-title"><div><span className="step">02</span><h3>未来6个月进销存滚动</h3></div><span className="formula">RMB吊牌万元 · 期初 + 到货 − 销售 = 期末</span></div>
          <div className="rolling-table data-roll">
            <div className="rolling-head"><span>月份</span><span>期初</span><span>预计到货</span><span>销售预测</span><span>期末</span><span>状态</span></div>
            {rolling.map((row) => (
              <div className="rolling-row" key={row.month}>
                <strong>{row.month}</strong><span>{fmt(row.opening)}</span><span className="arrival">+{fmt(row.arrivals)}</span><span className="sales">−{fmt(row.sales)}</span>
                <span className={row.closing < 0 ? "negative" : row.closing < overviewVelocityAmount / 10000 ? "low" : ""}>{fmt(row.closing, 1)}</span>
                <span>{row.closing < 0 ? "不足" : row.closing < overviewVelocityAmount / 10000 ? "偏低" : "可用"}</span>
              </div>
            ))}
          </div>
          {!sales || !inventory ? <div className="rolling-alert"><b>口径提示</b><span>当前为占位计算；补齐销售和库存后才形成真实滚动结果。OIH 默认放在第二个月到货。</span></div> : null}
        </section>
      </div>

      <p className="privacy-note">本机保存说明：清除浏览器网站数据、使用隐私模式或更换设备后不会自动带入；建议每次月末更新后导出一份备份。团队共享仍应使用统一的月结文件版本。</p>
    </section>
  );
}
