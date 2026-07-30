"use client";

import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type DataKind = "sales" | "inventory" | "oih";
type Row = Record<string, unknown>;

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
  const monthly: Record<string, number> = {};
  const sku: Record<string, number> = {};
  const countries: Record<string, number> = {};
  let quantity = 0;

  best.rows.forEach((row) => {
    const qty = quantityHeader ? numeric(row[quantityHeader]) : 0;
    quantity += qty;
    add(sku, skuHeader ? String(row[skuHeader] ?? "").trim() : "", qty);
    add(countries, countryHeader ? String(row[countryHeader] ?? "").trim() : "", qty);
    add(monthly, dateHeader ? monthOf(row[dateHeader]) : "", qty);
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

export default function DataHub() {
  const [datasets, setDatasets] = useState<Partial<Record<DataKind, Dataset>>>({});
  const [busy, setBusy] = useState<DataKind | null>(null);
  const [notice, setNotice] = useState("正在读取本机已保存的数据…");
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
