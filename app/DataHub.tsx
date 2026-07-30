"use client";

import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type DataKind = "sales" | "inventory" | "oih";
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
  quantity: number;
};

type Dataset = {
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
    skuCount: number;
    monthCount: number;
    countryCount: number;
    latestMonth: string;
  };
  monthly: Record<string, number>;
  sku: Record<string, number>;
  countries: Record<string, number>;
  details?: DetailRecord[];
};

const DB_NAME = "wilson-rolling-inventory";
const STORE = "datasets";
const kinds: Array<{ kind: DataKind; title: string; subtitle: string; accent: string }> = [
  { kind: "sales", title: "销售数据", subtitle: "观远 R01 / 零售销售", accent: "blue" },
  { kind: "inventory", title: "月末库存", subtitle: "BI 库存余额 / M01", accent: "green" },
  { kind: "oih", title: "OIH / 订单", subtitle: "在途、已下单及 VBR8", accent: "amber" },
];

const aliases = {
  sku: ["sku", "货号", "商品编码", "物料编码", "款号", "货品编码"],
  country: ["国家", "市场", "区域", "country", "market"],
  name: ["商品名称", "商品名", "品名", "货品名称", "产品名称", "sku名称"],
  brand: ["品牌", "品牌名称", "brand"],
  category: ["大类", "一级品类", "商品大类", "categoryl1", "category1"],
  middleCategory: ["中类", "二级品类", "商品中类", "categoryl2", "category2"],
  smallCategory: ["小类", "三级品类", "商品小类", "categoryl3", "category3", "品类"],
  series: ["系列", "商品系列", "产品系列", "series"],
  channel: ["渠道", "销售渠道", "店铺", "门店", "客户", "channel", "store"],
  date: ["日期", "月份", "年月", "销售日期", "库存月份", "快照日期", "date", "month"],
  sales: ["零售数量", "销售数量", "销量", "净销售数量", "销售件数", "数量", "qty", "quantity"],
  inventory: ["期末库存数量", "库存数量", "可用库存", "期末库存", "库存", "数量", "qty", "quantity"],
  oih: ["oih", "在途数量", "订单数量", "下单数量", "未到货数量", "open order", "数量", "qty", "quantity"],
};

function clean(value: unknown) {
  return String(value ?? "").trim().toLowerCase().replace(/[\s_\-\/（）()]/g, "");
}

function findHeader(headers: string[], options: string[]) {
  const normalized = options.map(clean);
  return headers.find((header) => normalized.some((option) => clean(header) === option || clean(header).includes(option)));
}

function numeric(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? "").replace(/,/g, "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function monthOf(value: unknown) {
  if (typeof value === "number" && value > 30000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? `${parsed.y}-${String(parsed.m).padStart(2, "0")}` : "";
  }
  const text = String(value ?? "").trim();
  const match = text.match(/(20\d{2})\D?([01]?\d)/);
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
      const quantityOptions = kind === "sales" ? aliases.sales : kind === "inventory" ? aliases.inventory : aliases.oih;
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
  const quantityOptions = kind === "sales" ? aliases.sales : kind === "inventory" ? aliases.inventory : aliases.oih;
  const quantityHeader = findHeader(best.headers, quantityOptions);
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
  const monthly: Record<string, number> = {};
  const sku: Record<string, number> = {};
  const countries: Record<string, number> = {};
  const detailMap = new Map<string, DetailRecord>();
  let quantity = 0;

  best.rows.forEach((row) => {
    const qty = quantityHeader ? numeric(row[quantityHeader]) : 0;
    quantity += qty;
    add(sku, skuHeader ? String(row[skuHeader] ?? "").trim() : "", qty);
    add(countries, countryHeader ? String(row[countryHeader] ?? "").trim() : "", qty);
    add(monthly, dateHeader ? monthOf(row[dateHeader]) : "", qty);
    const month = dateHeader ? monthOf(row[dateHeader]) : "";
    const skuValue = skuHeader ? String(row[skuHeader] ?? "").trim() : "";
    const country = countryHeader ? String(row[countryHeader] ?? "").trim() : "";
    const name = nameHeader ? String(row[nameHeader] ?? "").trim() : "";
    const brand = brandHeader ? String(row[brandHeader] ?? "").trim() : "";
    const category = categoryHeader ? String(row[categoryHeader] ?? "").trim() : "";
    const middleCategory = middleCategoryHeader ? String(row[middleCategoryHeader] ?? "").trim() : "";
    const smallCategory = smallCategoryHeader ? String(row[smallCategoryHeader] ?? "").trim() : "";
    const series = seriesHeader ? String(row[seriesHeader] ?? "").trim() : "";
    const channel = channelHeader ? String(row[channelHeader] ?? "").trim() : "";
    if (skuValue) {
      const key = [month, country, skuValue, name, brand, category, middleCategory, smallCategory, series, channel].join("¦");
      const current = detailMap.get(key);
      if (current) current.quantity += qty;
      else detailMap.set(key, { month, sku: skuValue, name, brand, country, category, middleCategory, smallCategory, series, channel, quantity: qty });
    }
  });
  const months = Object.keys(monthly).sort();
  const missing = [
    !quantityHeader && "数量",
    !skuHeader && "SKU",
    !dateHeader && "日期/月度",
  ].filter(Boolean);

  return {
    kind,
    fileName,
    uploadedAt: new Date().toISOString(),
    rowCount: best.rows.filter((row) => Object.values(row).some((value) => String(value).trim())).length,
    sheetName: best.sheetName,
    headers: best.headers,
    status: missing.length ? "mapping" : "ready",
    message: missing.length ? `已保存，但未识别：${missing.join("、")}` : "字段已自动识别，可用于滚动分析",
    summary: {
      quantity,
      skuCount: Object.keys(sku).length,
      monthCount: months.length,
      countryCount: Object.keys(countries).length,
      latestMonth: months.at(-1) ?? "",
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

function nextMonths(start: string, count = 6) {
  const base = start ? new Date(`${start}-01T00:00:00`) : new Date();
  return Array.from({ length: count }, (_, index) => {
    const date = new Date(base.getFullYear(), base.getMonth() + index + 1, 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  });
}

export default function DataHub({ view = "data" }: { view?: "data" | "analytics" }) {
  const [datasets, setDatasets] = useState<Partial<Record<DataKind, Dataset>>>({});
  const [busy, setBusy] = useState<DataKind | null>(null);
  const [notice, setNotice] = useState("正在读取本机已保存的数据…");
  const [analysisMonth, setAnalysisMonth] = useState("");
  const [countryFilter, setCountryFilter] = useState("全部");
  const [categoryFilter, setCategoryFilter] = useState("全部");
  const [brandFilter, setBrandFilter] = useState("全部");
  const [middleCategoryFilter, setMiddleCategoryFilter] = useState("全部");
  const [smallCategoryFilter, setSmallCategoryFilter] = useState("全部");
  const [seriesFilter, setSeriesFilter] = useState("全部");
  const [skuSearch, setSkuSearch] = useState("");
  const [riskFilter, setRiskFilter] = useState("全部");
  const inputs = useRef<Partial<Record<DataKind, HTMLInputElement | null>>>({});

  useEffect(() => {
    getAll()
      .then((rows) => {
        setDatasets(Object.fromEntries(rows.map((row) => [row.kind, row])));
        setNotice(rows.length ? `已恢复 ${rows.length} 类数据，无需再次上传` : "尚未上传真实数据，可从三个数据源开始");
      })
      .catch(() => setNotice("浏览器存储不可用，请检查隐私模式或网站存储权限"));
  }, []);

  const sales = datasets.sales;
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
  const detailReady = Boolean(sales?.details?.length && inventory?.details?.length);
  const availableMonths = Array.from(new Set((sales?.details ?? []).map((row) => row.month).filter(Boolean))).sort();
  const selectedMonth = analysisMonth || availableMonths.at(-1) || "";
  const recentAnalysisMonths = selectedMonth ? availableMonths.filter((month) => month <= selectedMonth).slice(-3) : [];
  const latestInventoryMonth = Array.from(new Set((inventory?.details ?? []).map((row) => row.month).filter(Boolean))).sort().at(-1) ?? "";
  const filterCountries = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.country).filter(Boolean))).sort();
  const filterCategories = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.category).filter(Boolean))).sort();
  const filterBrands = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.brand).filter(Boolean))).sort();
  const filterMiddleCategories = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].filter((row) => categoryFilter === "全部" || row.category === categoryFilter).map((row) => row.middleCategory).filter(Boolean))).sort();
  const filterSmallCategories = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].filter((row) => (categoryFilter === "全部" || row.category === categoryFilter) && (middleCategoryFilter === "全部" || row.middleCategory === middleCategoryFilter)).map((row) => row.smallCategory).filter(Boolean))).sort();
  const filterSeries = Array.from(new Set([...(sales?.details ?? []), ...(inventory?.details ?? [])].map((row) => row.series).filter(Boolean))).sort();

  const skuAnalysis = useMemo(() => {
    type SkuLine = { key: string; sku: string; name: string; brand: string; country: string; category: string; middleCategory: string; smallCategory: string; series: string; channel: string; stock: number; sales: number; sales3m: number; oih: number; ratio: number; weeks: number; status: string };
    const lines = new Map<string, SkuLine>();
    const ensure = (row: DetailRecord) => {
      const key = `${row.country}¦${row.sku}`;
      if (!lines.has(key)) lines.set(key, { key, sku: row.sku, name: row.name, brand: row.brand, country: row.country, category: row.category, middleCategory: row.middleCategory, smallCategory: row.smallCategory, series: row.series, channel: row.channel, stock: 0, sales: 0, sales3m: 0, oih: 0, ratio: 0, weeks: 0, status: "" });
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
    (inventory?.details ?? []).filter((row) => !latestInventoryMonth || !row.month || row.month === latestInventoryMonth).forEach((row) => { ensure(row).stock += row.quantity; });
    (sales?.details ?? []).filter((row) => recentAnalysisMonths.includes(row.month)).forEach((row) => {
      const line = ensure(row);
      line.sales3m += row.quantity / Math.max(recentAnalysisMonths.length, 1);
      if (row.month === selectedMonth) line.sales += row.quantity;
    });
    (oih?.details ?? []).forEach((row) => { ensure(row).oih += row.quantity; });
    return Array.from(lines.values()).map((line) => {
      line.ratio = line.sales3m > 0 ? line.stock / line.sales3m : line.stock > 0 ? 999 : 0;
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
  }, [sales?.details, inventory?.details, oih?.details, latestInventoryMonth, selectedMonth, recentAnalysisMonths.join("|"), countryFilter, brandFilter, categoryFilter, middleCategoryFilter, smallCategoryFilter, seriesFilter, riskFilter, skuSearch]);

  const analyticStock = skuAnalysis.reduce((sum, row) => sum + row.stock, 0);
  const analyticSales = skuAnalysis.reduce((sum, row) => sum + row.sales, 0);
  const analyticVelocity = skuAnalysis.reduce((sum, row) => sum + row.sales3m, 0);
  const analyticRatio = analyticVelocity ? analyticStock / analyticVelocity : 0;
  const stockedSku = skuAnalysis.filter((row) => row.stock > 0).length;
  const movingSku = skuAnalysis.filter((row) => row.sales3m > 0).length;
  const zeroMovingStock = skuAnalysis.filter((row) => row.status === "无动销").reduce((sum, row) => sum + row.stock, 0);
  const healthyRate = stockedSku ? skuAnalysis.filter((row) => row.status === "健康").length / stockedSku : 0;
  const dimensionSummary = (field: "brand" | "category" | "middleCategory" | "smallCategory" | "series" | "country") => {
    const grouped = new Map<string, { name: string; stock: number; sales: number }>();
    skuAnalysis.forEach((row) => {
      const name = row[field] || "未识别";
      const current = grouped.get(name) ?? { name, stock: 0, sales: 0 };
      current.stock += row.stock;
      current.sales += row.sales3m;
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
    let opening = inventoryQty;
    const forecast = monthlyVelocity;
    return months.map((month, index) => {
      const arrivals = index === 1 ? oihQty : 0;
      const closing = opening + arrivals - forecast;
      const row = { month, opening, arrivals, sales: forecast, closing };
      opening = closing;
      return row;
    });
  }, [inventoryQty, monthlyVelocity, oihQty, sales?.summary.latestMonth, inventory?.summary.latestMonth]);

  async function upload(kind: DataKind, event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setBusy(kind);
    setNotice(`正在解析 ${file.name}…`);
    try {
      const parsed = parseWorkbook(await file.arrayBuffer(), kind, file.name);
      await putOne(parsed);
      setDatasets((current) => ({ ...current, [kind]: parsed }));
      setNotice(`${file.name} 已保存；下次打开会自动恢复`);
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

  const conclusion = !inventory || !sales
    ? "先补齐销售与月末库存，才能形成可靠的库存周数和滚动结论。"
    : woi > 26
      ? `当前库存约 ${woi.toFixed(1)} 周，偏高；建议先消化库存，并将新订单拆分到货。`
      : rolling.some((row) => row.closing < 0)
        ? "滚动预测期内将出现缺货；建议按缺口补单，但不要一次性提前全部到货。"
        : `当前库存约 ${woi.toFixed(1)} 周，滚动期内库存可控；订单仍需按国家与 SKU 继续审核。`;

  if (view === "analytics") {
    return (
      <section className="data-hub analytics-view">
        <div className="page-heading">
          <div><p className="eyebrow blue">SALES & INVENTORY DIAGNOSTICS</p><h2>销售与库存经营看板</h2><p>按月份、国家、品类和SKU下钻，识别无动销、高库存与缺货风险。上传数据仍保存在当前浏览器。</p></div>
          <div className="local-badge"><i /> 同一设备自动记住</div>
        </div>
        <div className="analytics-filterbar">
          <label>分析月份<select value={selectedMonth} onChange={(event) => setAnalysisMonth(event.target.value)}><option value="">最新月份</option>{availableMonths.map((month) => <option key={month}>{month}</option>)}</select></label>
          <label>品牌<select value={brandFilter} onChange={(event) => setBrandFilter(event.target.value)}><option>全部</option>{filterBrands.map((brand) => <option key={brand}>{brand}</option>)}</select></label>
          <label>国家<select value={countryFilter} onChange={(event) => setCountryFilter(event.target.value)}><option>全部</option>{filterCountries.map((country) => <option key={country}>{country}</option>)}</select></label>
          <label>大类<select value={categoryFilter} onChange={(event) => { setCategoryFilter(event.target.value); setMiddleCategoryFilter("全部"); setSmallCategoryFilter("全部"); }}><option>全部</option>{filterCategories.map((category) => <option key={category}>{category}</option>)}</select></label>
          <label>中类<select value={middleCategoryFilter} onChange={(event) => { setMiddleCategoryFilter(event.target.value); setSmallCategoryFilter("全部"); }}><option>全部</option>{filterMiddleCategories.map((category) => <option key={category}>{category}</option>)}</select></label>
          <label>小类<select value={smallCategoryFilter} onChange={(event) => setSmallCategoryFilter(event.target.value)}><option>全部</option>{filterSmallCategories.map((category) => <option key={category}>{category}</option>)}</select></label>
          <label>系列<select value={seriesFilter} onChange={(event) => setSeriesFilter(event.target.value)}><option>全部</option>{filterSeries.map((series) => <option key={series}>{series}</option>)}</select></label>
          <label>库存状态<select value={riskFilter} onChange={(event) => setRiskFilter(event.target.value)}>{["全部", "健康", "高库存", "无动销", "缺货风险"].map((status) => <option key={status}>{status}</option>)}</select></label>
          <label className="search-filter">SKU / 商品<input value={skuSearch} onChange={(event) => setSkuSearch(event.target.value)} placeholder="输入货号或品名" /></label>
        </div>
        {!detailReady && <div className="detail-upgrade-note"><strong>需要重新上传销售与库存文件</strong><span>旧版只保存总量汇总；新版上传后会保留SKU、国家、品类、渠道和月份维度，才能生成细颗粒度分析。</span></div>}
        <div className="summary-grid analytics-kpis">
          <article><span>期末库存</span><strong>{detailReady ? fmt(analyticStock) : "—"} <em>件</em></strong><small>{latestInventoryMonth || "等待库存月份"}</small></article>
          <article><span>当月销售</span><strong>{detailReady ? fmt(analyticSales) : "—"} <em>件</em></strong><small>{selectedMonth || "等待销售月份"}</small></article>
          <article><span>库销比</span><strong className={analyticRatio > 6 ? "red" : analyticRatio > 4 ? "amber" : ""}>{detailReady ? analyticRatio.toFixed(1) : "—"} <em>月</em></strong><small>库存 ÷ 近3月月均销量</small></article>
          <article><span>动销率</span><strong>{detailReady && stockedSku ? `${((movingSku / stockedSku) * 100).toFixed(1)}%` : "—"}</strong><small>有销量SKU ÷ 有库存SKU</small></article>
          <article><span>库存宽度</span><strong>{detailReady ? fmt(stockedSku) : "—"} <em>SKU</em></strong><small>当前有库存的SKU数量</small></article>
          <article><span>平均深度</span><strong>{detailReady && stockedSku ? fmt(analyticStock / stockedSku, 1) : "—"} <em>件/SKU</em></strong><small>库存数量 ÷ 有库存SKU</small></article>
          <article><span>无动销库存</span><strong className="red">{detailReady ? fmt(zeroMovingStock) : "—"} <em>件</em></strong><small>近3月无销售但仍有库存</small></article>
          <article><span>健康SKU占比</span><strong>{detailReady ? `${(healthyRate * 100).toFixed(1)}%` : "—"}</strong><small>库销比1–6个月且持续动销</small></article>
        </div>
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
            const data = rows as Array<{ name: string; stock: number; sales: number }>;
            const maxStock = Math.max(...data.map((row) => row.stock), 1);
            return <section className="panel dimension-panel" key={title as string}><div className="panel-title"><div><h3>{title as string}</h3></div><span className="legend"><span className="line-key actual" />库存 <span className="line-key target" />月均销售</span></div>
              <div className="dimension-list">{data.length ? data.map((row) => <div key={row.name}><strong>{row.name}</strong><div className="dual-bar"><i style={{ width: `${row.stock / maxStock * 100}%` }} /><b style={{ width: `${Math.min(100, row.sales / maxStock * 100)}%` }} /></div><span>{fmt(row.stock)} / {fmt(row.sales, 1)}</span></div>) : <p>等待维度数据</p>}</div>
            </section>;
          })}
        </div>
        <section className="panel sku-diagnostic">
          <div className="panel-title"><div><span className="step">03</span><h3>SKU经营诊断明细</h3></div><span className="unit">{fmt(skuAnalysis.length)} 行 · 点击上方筛选</span></div>
          <div className="table-wrap"><table><thead><tr><th>SKU / 商品</th><th>品牌</th><th>国家</th><th>大/中/小类</th><th>系列 / 渠道</th><th>期末库存</th><th>当月销售</th><th>近3月月均</th><th>库销比</th><th>库存周数</th><th>OIH</th><th>诊断</th></tr></thead>
          <tbody>{skuAnalysis.slice(0, 500).map((row) => <tr key={row.key}><td><strong>{row.sku}</strong><small>{row.name || "未识别品名"}</small></td><td>{row.brand || "未识别"}</td><td>{row.country || "未识别"}</td><td><strong>{row.category || "—"} / {row.middleCategory || "—"}</strong><small>{row.smallCategory || "未识别小类"}</small></td><td><strong>{row.series || "未识别系列"}</strong><small>{row.channel || "未识别渠道"}</small></td><td><strong>{fmt(row.stock)}</strong></td><td>{fmt(row.sales)}</td><td>{fmt(row.sales3m, 1)}</td><td>{row.ratio >= 999 ? "无销量" : `${row.ratio.toFixed(1)}月`}</td><td>{row.weeks >= 999 ? "无销量" : `${row.weeks.toFixed(1)}周`}</td><td>{fmt(row.oih)}</td><td><span className={`diagnostic-tag ${row.status}`}>{row.status}</span></td></tr>)}</tbody></table></div>
          {skuAnalysis.length > 500 && <p className="table-limit">当前显示库存最高的500行，请使用筛选缩小范围。</p>}
        </section>
        <p className="privacy-note">计算口径：库存取最新可识别库存月份；当月销售取所选月份；库销比和库存周数使用截至所选月份的最近3个月月均销量。无销量但有库存归类为“无动销”。</p>
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
                  <div className="upload-stats"><span>数量 <b>{fmt(data.summary.quantity)}</b></span><span>SKU <b>{fmt(data.summary.skuCount)}</b></span><span>月份 <b>{fmt(data.summary.monthCount)}</b></span></div>
                  <small className={data.status === "ready" ? "parse-ok" : "parse-warning"}>{data.message}</small>
                </>
              ) : <p className="empty-copy">支持 .xlsx、.xls、.csv；会自动寻找前25行中的表头并识别数量、SKU、国家和月份。</p>}
              <div className="upload-actions">
                <input ref={(node) => { inputs.current[source.kind] = node; }} hidden type="file" accept=".xlsx,.xls,.csv" onChange={(event) => upload(source.kind, event)} />
                <button className="primary small" disabled={busy !== null} onClick={() => inputs.current[source.kind]?.click()}>{busy === source.kind ? "正在解析…" : data ? "覆盖上传" : "选择文件"}</button>
                {data && <button className="danger-link" onClick={() => remove(source.kind)}>删除</button>}
              </div>
            </article>
          );
        })}
      </div>

      <div className="summary-grid data-kpis">
        <article><span>库存宽度</span><strong>{activeSku ? fmt(activeSku) : "—"} <em>SKU</em></strong><small>销售、库存及 OIH 去重后 SKU 数</small></article>
        <article><span>平均库存深度</span><strong>{inventory ? fmt(depth, 1) : "—"} <em>件/SKU</em></strong><small>期末库存 ÷ 有库存 SKU</small></article>
        <article><span>最近月均销量</span><strong>{sales ? fmt(monthlyVelocity, 1) : "—"} <em>件/月</em></strong><small>最近3个可识别月份平均</small></article>
        <article><span>库存周数 WOI</span><strong className={woi > 26 ? "red" : woi > 20 ? "amber" : ""}>{sales && inventory ? `${woi.toFixed(1)}周` : "—"}</strong><small>期末库存 ÷ 月均销量 × 4.33</small></article>
      </div>

      <section className="data-conclusion">
        <div><span>当前系统结论</span><h3>{conclusion}</h3></div>
        <strong>{sales && inventory ? (woi > 26 ? "建议控单" : "进入SKU审核") : "数据待补齐"}</strong>
      </section>

      <div className="data-analysis-grid">
        <section className="panel">
          <div className="panel-title"><div><span className="step">01</span><h3>近月销售趋势</h3></div><span className="unit">实际数量</span></div>
          {salesMonths.length ? (
            <div className="mini-trend">
              {salesMonths.slice(-12).map(([month, value]) => {
                const max = Math.max(...salesMonths.slice(-12).map(([, number]) => number), 1);
                return <div key={month}><span>{month.slice(2)}</span><i style={{ height: `${Math.max(5, (value / max) * 100)}%` }} /><b>{fmt(value)}</b></div>;
              })}
            </div>
          ) : <div className="empty-panel">上传销售文件后显示最近12个月趋势</div>}
        </section>

        <section className="panel">
          <div className="panel-title"><div><span className="step">02</span><h3>未来6个月进销存滚动</h3></div><span className="formula">期初 + 到货 − 销售 = 期末</span></div>
          <div className="rolling-table data-roll">
            <div className="rolling-head"><span>月份</span><span>期初</span><span>预计到货</span><span>销售预测</span><span>期末</span><span>状态</span></div>
            {rolling.map((row) => (
              <div className="rolling-row" key={row.month}>
                <strong>{row.month}</strong><span>{fmt(row.opening)}</span><span className="arrival">+{fmt(row.arrivals)}</span><span className="sales">−{fmt(row.sales)}</span>
                <span className={row.closing < 0 ? "negative" : row.closing < monthlyVelocity ? "low" : ""}>{fmt(row.closing)}</span>
                <span>{row.closing < 0 ? "缺货" : row.closing < monthlyVelocity ? "偏低" : "可用"}</span>
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
