"use client";

import { useMemo, useState } from "react";
import DataHub from "./DataHub";

type OrderType = "客户锁单" | "常规补货" | "新品首单" | "电商补货" | "国家仓备货";
type Status = "通过" | "有条件通过" | "需调整";

type StockItem = {
  sku: string;
  name: string;
  country: string;
  current: number;
  oih: number;
  sales13m: number;
  sales3m: number;
  orderQty: number;
  unitCost: number;
  orderType: OrderType;
  paymentMonth: string;
  customerEvidence: boolean;
  status: Status;
  reasons: string[];
};

const seedItems: StockItem[] = [
  {
    sku: "DEMO-001",
    name: "示例客户锁单商品",
    country: "示例市场A",
    current: 12,
    oih: 20,
    sales13m: 72,
    sales3m: 18,
    orderQty: 24,
    unitCost: 50,
    orderType: "客户锁单",
    paymentMonth: "2026-10",
    customerEvidence: true,
    status: "有条件通过",
    reasons: ["虚拟示例：客户订单已锁定", "需确认不可取消条款与回款日期", "本次订单尚未进入OIH，按增量影响审核"],
  },
  {
    sku: "DEMO-002",
    name: "示例稳定补货商品",
    country: "示例市场B",
    current: 35,
    oih: 12,
    sales13m: 104,
    sales3m: 27,
    orderQty: 18,
    unitCost: 42,
    orderType: "客户锁单",
    paymentMonth: "2026-10",
    customerEvidence: true,
    status: "有条件通过",
    reasons: ["虚拟示例：订单属性为客户锁单", "需核对客户最终数量", "通过条件：信用额度及付款条款确认"],
  },
  {
    sku: "DEMO-003",
    name: "示例新品首单商品",
    country: "示例市场C",
    current: 0,
    oih: 168,
    sales13m: 0,
    sales3m: 0,
    orderQty: 48,
    unitCost: 60,
    orderType: "新品首单",
    paymentMonth: "2026-11",
    customerEvidence: false,
    status: "需调整",
    reasons: ["虚拟示例：新品尚无零售销售", "现有OIH到货后库存已偏高，本次订单将进一步增加库存", "缺少门店首铺计划或客户锁单，建议暂缓"],
  },
  {
    sku: "DEMO-004",
    name: "示例常规补货商品",
    country: "示例市场D",
    current: 42,
    oih: 24,
    sales13m: 61,
    sales3m: 21,
    orderQty: 24,
    unitCost: 48,
    orderType: "常规补货",
    paymentMonth: "2026-10",
    customerEvidence: false,
    status: "通过",
    reasons: ["近3个月月均销量7支，动销稳定", "下单后约12.9个月库存，低于常青款上限", "订单金额处于国家月度额度内"],
  },
];

const statusMeta: Record<Status, { label: string; className: string }> = {
  通过: { label: "建议通过", className: "good" },
  有条件通过: { label: "有条件通过", className: "warn" },
  需调整: { label: "建议调整", className: "bad" },
};

function money(value: number) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(value);
}

function weeksOfInventory(item: StockItem) {
  const monthlySales = item.sales3m > 0 ? item.sales3m / 3 : 0;
  if (!monthlySales) return null;
  return ((item.current + item.oih + item.orderQty) / monthlySales) * 4.33;
}

export default function Home() {
  const [items, setItems] = useState(seedItems);
  const [selectedSku, setSelectedSku] = useState(seedItems[2].sku);
  const [decision, setDecision] = useState("待财务决策");
  const [activeTab, setActiveTab] = useState<"data" | "sales" | "inventory" | "overview" | "cash" | "review" | "rules">("data");
  const [draft, setDraft] = useState({
    sku: "DEMO-003",
    country: "示例市场C",
    orderQty: 48,
    orderType: "新品首单" as OrderType,
    paymentMonth: "2026-11",
  });

  const selected = items.find((item) => item.sku === selectedSku) ?? items[0];
  const totalValue = items.reduce((sum, item) => sum + item.orderQty * item.unitCost, 0);
  const risky = items.filter((item) => item.status === "需调整").length;
  const conditional = items.filter((item) => item.status === "有条件通过").length;
  const totalQty = items.reduce((sum, item) => sum + item.orderQty, 0);
  const holdItems = items.filter((item) => item.status === "需调整");
  const holdQty = holdItems.reduce((sum, item) => sum + item.orderQty, 0);
  const holdValue = holdItems.reduce((sum, item) => sum + item.orderQty * item.unitCost, 0);
  const recommendedQty = totalQty - holdQty;
  const recommendedValue = totalValue - holdValue;

  const payments = useMemo(() => {
    const grouped = new Map<string, number>();
    items.forEach((item) => grouped.set(item.paymentMonth, (grouped.get(item.paymentMonth) ?? 0) + item.orderQty * item.unitCost));
    return Array.from(grouped.entries()).sort(([a], [b]) => a.localeCompare(b));
  }, [items]);

  const salesTrend = [
    ["2026-02", 286, 310],
    ["2026-03", 342, 330],
    ["2026-04", 318, 350],
    ["2026-05", 389, 370],
    ["2026-06", 421, 410],
    ["2026-07", 447, 440],
  ] as const;

  const rollingPlan = [
    { month: "2026-08", opening: 925, arrivals: 286, sales: 472, closing: 739, payment: 498 },
    { month: "2026-09", opening: 739, arrivals: 164, sales: 495, closing: 408, payment: 441 },
    { month: "2026-10", opening: 408, arrivals: 236, sales: 518, closing: 126, payment: 490 },
    { month: "2026-11", opening: 126, arrivals: 318, sales: 462, closing: -18, payment: 246 },
    { month: "2026-12", opening: -18, arrivals: 274, sales: 431, closing: -175, payment: 225 },
    { month: "2027-01", opening: -175, arrivals: 482, sales: 408, closing: -101, payment: 338 },
  ];

  function applyDraft() {
    const existing = items.find((item) => item.sku === draft.sku);
    if (!existing) return;
    setItems((current) =>
      current.map((item) =>
        item.sku === draft.sku
          ? { ...item, country: draft.country, orderQty: Number(draft.orderQty), orderType: draft.orderType, paymentMonth: draft.paymentMonth }
          : item,
      ),
    );
    setSelectedSku(draft.sku);
  }

  return (
    <main>
      <header className="topbar">
        <div className="brandmark">W</div>
        <div>
          <p className="eyebrow">MONTHLY BUSINESS ANALYTICS</p>
          <h1>Sales & Inventory Management Center</h1>
        </div>
        <div className="header-right">
          <span className="data-pill"><i /> 数据更新至 2026-07-28</span>
          <button className="ghost">导出审核记录</button>
          <div className="avatar">QY</div>
        </div>
      </header>

      <nav className="tabs" aria-label="页面导航">
        <button className={activeTab === "data" ? "active" : ""} onClick={() => setActiveTab("data")}>数据中心</button>
        <button className={activeTab === "sales" ? "active" : ""} onClick={() => setActiveTab("sales")}>销售看板</button>
        <button className={activeTab === "inventory" ? "active" : ""} onClick={() => setActiveTab("inventory")}>库存看板</button>
        <button className={activeTab === "overview" ? "active" : ""} onClick={() => setActiveTab("overview")}>进销存总览</button>
        <button className={activeTab === "cash" ? "active" : ""} onClick={() => setActiveTab("cash")}>现金流预测</button>
        <button className={activeTab === "review" ? "active" : ""} onClick={() => setActiveTab("review")}>下单审核</button>
        <button className={activeTab === "rules" ? "active" : ""} onClick={() => setActiveTab("rules")}>规则与口径</button>
        <span className="source-note">零售 · 批发 · 库存 · OIH · 订单</span>
      </nav>

      <section className="workspace">
        {activeTab === "data" && <DataHub />}
        {activeTab === "sales" && <DataHub view="sales" />}
        {activeTab === "inventory" && <DataHub view="inventory" />}
        {activeTab === "overview" && <DataHub view="analytics" />}
        {activeTab === "cash" && <DataHub view="cash" />}
        {activeTab === "review" && (
          <>
            <div className="page-heading">
              <div>
                <p className="eyebrow blue">ORDER REVIEW / VBR8</p>
                <h2>本次下单审核</h2>
                <p>所有判断由公开规则生成。点击任一SKU，可查看数据来源、计算过程和触发原因。</p>
              </div>
              <div className="decision-box">
                <span>当前决策</span>
                <strong>{decision}</strong>
              </div>
            </div>

            <div className="summary-grid">
              <article><span>本次下单金额</span><strong>{money(totalValue)}</strong><small>4个样本SKU · VBR8</small></article>
              <article><span>建议调整</span><strong className="red">{risky} <em>SKU</em></strong><small>触发硬性风险规则</small></article>
              <article><span>有条件通过</span><strong className="amber">{conditional} <em>SKU</em></strong><small>需补充客户或付款证据</small></article>
              <article><span>预计付款峰值</span><strong>{money(Math.max(...payments.map(([, value]) => value)))}</strong><small>{payments.sort((a, b) => b[1] - a[1])[0]?.[0]} 预计付款</small></article>
            </div>

            <section className="executive-review" aria-label="本次订单审核总览">
              <div className="executive-verdict">
                <span className="verdict-label">本单综合结论</span>
                <div className="verdict-title">
                  <i className="verdict-icon">!</i>
                  <div>
                    <h3>建议有条件批准，不建议整单直接放行</h3>
                    <p>批准有明确需求依据的部分，同时暂缓高库存且缺少铺货证据的新品订单。</p>
                  </div>
                </div>
                <div className="verdict-numbers">
                  <div><span>提交数量</span><strong>{totalQty}</strong><small>支</small></div>
                  <b>→</b>
                  <div className="recommended"><span>建议批准</span><strong>{recommendedQty}</strong><small>支 · {money(recommendedValue)}</small></div>
                  <div className="held"><span>建议暂缓</span><strong>{holdQty}</strong><small>支 · {money(holdValue)}</small></div>
                </div>
              </div>

              <div className="executive-reasons">
                <div className="executive-heading">
                  <h3>财务审核意见</h3>
                  <span>基于库存、销售、OIH及现金影响</span>
                </div>
                <ol>
                  <li className="positive"><b>可批准</b><span>{recommendedQty}支具备客户锁单或稳定动销依据，但其中客户锁单须补齐不可取消条款与回款日期。</span></li>
                  <li className="negative"><b>需暂缓</b><span>{holdQty}支新品在现有OIH到货后仍缺少销售支撑，补充门店铺货计划或客户订单后再审。</span></li>
                  <li className="cash"><b>现金安排</b><span>付款集中在10月；本批金额尚可控，但应与其他到货付款合并检查月度额度后放行。</span></li>
                </ol>
                <div className="next-action"><strong>建议动作</strong><span>拆单审批：先释放{recommendedQty}支，冻结{holdQty}支，并将资料补充设为放行条件。</span></div>
              </div>
            </section>

            <div className="main-grid">
              <section className="panel order-input">
                <div className="panel-title">
                  <div><span className="step">01</span><h3>业务订单输入</h3></div>
                  <span className="saved">表单已校验</span>
                </div>
                <div className="form-grid">
                  <label>货号 / SKU
                    <select value={draft.sku} onChange={(e) => setDraft({ ...draft, sku: e.target.value })}>
                      {items.map((item) => <option key={item.sku}>{item.sku}</option>)}
                    </select>
                  </label>
                  <label>国家
                    <select value={draft.country} onChange={(e) => setDraft({ ...draft, country: e.target.value })}>
                      {["示例市场A", "示例市场B", "示例市场C", "示例市场D"].map((country) => <option key={country}>{country}</option>)}
                    </select>
                  </label>
                  <label>订单类型
                    <select value={draft.orderType} onChange={(e) => setDraft({ ...draft, orderType: e.target.value as OrderType })}>
                      {["客户锁单", "常规补货", "新品首单", "电商补货", "国家仓备货"].map((type) => <option key={type}>{type}</option>)}
                    </select>
                  </label>
                  <label>本次下单数量
                    <input type="number" value={draft.orderQty} onChange={(e) => setDraft({ ...draft, orderQty: Number(e.target.value) })} />
                  </label>
                  <label>预计付款月份
                    <input type="month" value={draft.paymentMonth} onChange={(e) => setDraft({ ...draft, paymentMonth: e.target.value })} />
                  </label>
                  <button className="primary" onClick={applyDraft}>重新运行审核</button>
                </div>
              </section>

              <section className="panel cash-panel">
                <div className="panel-title"><div><span className="step">02</span><h3>付款月份影响</h3></div><span className="unit">USD</span></div>
                <div className="cash-chart">
                  {payments.map(([month, value]) => (
                    <div className="bar-row" key={month}>
                      <span>{month}</span>
                      <div><i style={{ width: `${Math.max(12, (value / Math.max(...payments.map(([, v]) => v))) * 100)}%` }} /></div>
                      <strong>{money(value)}</strong>
                    </div>
                  ))}
                </div>
                <p className="cash-note">现金规则：若单月付款超过月度额度10%，自动升级至财务负责人审批。</p>
              </section>
            </div>

            <div className="rolling-grid">
              <section className="panel sales-panel">
                <div className="panel-title">
                  <div><span className="step">03</span><h3>月度销售趋势</h3></div>
                  <div className="legend"><span className="line-key actual" />实际销售 <span className="line-key target" />目标</div>
                </div>
                <div className="sales-kpis">
                  <div><span>近3月销售</span><strong>1,257</strong><small>件 / 支</small></div>
                  <div><span>较前3月</span><strong className="green-text">+25.8%</strong><small>销售速度提升</small></div>
                  <div><span>7月达成率</span><strong>101.6%</strong><small>实际447 / 目标440</small></div>
                </div>
                <div className="sales-chart">
                  {salesTrend.map(([month, actual, target]) => (
                    <div className="sales-month" key={month}>
                      <div className="sales-bars">
                        <i className="target-bar" style={{ height: `${(target / 500) * 100}%` }} />
                        <i className="actual-bar" style={{ height: `${(actual / 500) * 100}%` }} />
                      </div>
                      <strong>{actual}</strong>
                      <span>{month.slice(5)}月</span>
                    </div>
                  ))}
                </div>
                <p className="chart-footnote">销售口径：R01不含赠品零售数量；批发销售在正式版中单列，避免与零售重复。</p>
              </section>

              <section className="panel rolling-panel">
                <div className="panel-title">
                  <div><span className="step">04</span><h3>未来6个月进销存＋现金滚动</h3></div>
                  <span className="formula">期初 + 到货 − 销售 = 期末</span>
                </div>
                <div className="rolling-table">
                  <div className="rolling-head"><span>月份</span><span>期初库存</span><span>预计到货</span><span>销售预测</span><span>期末库存</span><span>付款 USD'000</span></div>
                  {rollingPlan.map((row) => (
                    <div className="rolling-row" key={row.month}>
                      <strong>{row.month}</strong>
                      <span>{row.opening}</span>
                      <span className="arrival">+{row.arrivals}</span>
                      <span className="sales">−{row.sales}</span>
                      <span className={row.closing < 0 ? "negative" : row.closing < 150 ? "low" : ""}>{row.closing}</span>
                      <span className={row.payment >= 450 ? "cash-risk" : ""}>{row.payment}</span>
                    </div>
                  ))}
                </div>
                <div className="rolling-alert"><b>滚动预警</b><span>11月开始出现潜在缺货，但8–10月付款集中。建议调整到货节奏，而不是简单扩大总订单。</span></div>
              </section>
            </div>

            <section className="panel audit-table">
              <div className="panel-title">
                <div><span className="step">05</span><h3>SKU审核结果</h3></div>
                <div className="legend"><span className="dot good" />通过 <span className="dot warn" />有条件 <span className="dot bad" />调整</div>
              </div>
              <div className="table-wrap">
                <table>
                  <thead><tr><th>SKU / 商品</th><th>国家</th><th>订单类型</th><th>当前库存</th><th>OIH</th><th>近3月销售</th><th>本次下单</th><th>下单后WOI</th><th>审核结论</th></tr></thead>
                  <tbody>
                    {items.map((item) => {
                      const woi = weeksOfInventory(item);
                      return (
                        <tr className={selectedSku === item.sku ? "selected" : ""} key={item.sku} onClick={() => setSelectedSku(item.sku)}>
                          <td><strong>{item.sku}</strong><small>{item.name}</small></td>
                          <td>{item.country}</td><td><span className="type-pill">{item.orderType}</span></td>
                          <td>{item.current}</td><td>{item.oih}</td><td>{item.sales3m}</td><td><strong>{item.orderQty}</strong></td>
                          <td>{woi ? `${woi.toFixed(1)}周` : "新品/无基数"}</td>
                          <td><span className={`status ${statusMeta[item.status].className}`}>{statusMeta[item.status].label}</span></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>

            <div className="detail-grid">
              <section className="panel evidence">
                <div className="panel-title"><div><span className="step">06</span><h3>{selected.sku} 审核解释</h3></div><span className={`status ${statusMeta[selected.status].className}`}>{statusMeta[selected.status].label}</span></div>
                <div className="stock-flow">
                  <div><span>当前库存</span><strong>{selected.current}</strong></div><b>+</b>
                  <div><span>现有OIH</span><strong>{selected.oih}</strong></div><b>+</b>
                  <div><span>本次下单</span><strong>{selected.orderQty}</strong></div><b>=</b>
                  <div className="result"><span>未来可用库存</span><strong>{selected.current + selected.oih + selected.orderQty}</strong></div>
                </div>
                <ul>
                  {selected.reasons.map((reason, index) => <li key={reason}><span>{index + 1}</span>{reason}</li>)}
                </ul>
                <div className="source-line"><span>数据追溯</span> M01库存 2606 · OIH 2026-07-01 · R01销售 2025-07至2026-07 · VBR8订单</div>
              </section>

              <section className="panel rule-panel">
                <div className="panel-title"><div><span className="step">07</span><h3>本单适用规则</h3></div><button onClick={() => setActiveTab("rules")} className="text-button">查看全部规则 →</button></div>
                <div className="rule-row"><span>订单类型优先</span><strong>{selected.orderType}</strong><i className="pass">已识别</i></div>
                <div className="rule-row"><span>增量口径检查</span><strong>现有OIH + 本次拟下订单</strong><i className="pass">无重复</i></div>
                <div className="rule-row"><span>库存周数规则</span><strong>{weeksOfInventory(selected)?.toFixed(1) ?? "新品豁免"}</strong><i className="pass">规则公开</i></div>
                <div className="rule-row"><span>客户证据</span><strong>{selected.customerEvidence ? "已提供" : "未提供"}</strong><i className={selected.customerEvidence ? "pass" : "attention"}>{selected.customerEvidence ? "通过" : "待补充"}</i></div>
              </section>
            </div>

            <section className="decision-footer">
              <div><p className="eyebrow">FINAL DECISION</p><h3>财务最终决策</h3><span>决策将连同规则命中记录留痕，业务可查看完整原因。</span></div>
              <div className="decision-actions">
                <button onClick={() => setDecision("退回业务调整")}>退回调整</button>
                <button onClick={() => setDecision("有条件批准")}>有条件批准</button>
                <button className="approve" onClick={() => setDecision("批准下单")}>批准下单</button>
              </div>
            </section>
          </>
        )}

        {false && (
          <section className="alternate">
            <div className="page-heading"><div><p className="eyebrow blue">INVENTORY HEALTH</p><h2>库存健康总览</h2><p>从品牌、国家、SKU和货季四个层级查看库存结果。</p></div></div>
            <div className="summary-grid inventory-cards">
              <article><span>Wilson期末库存</span><strong>924,567</strong><small>2606月末实际</small></article>
              <article><span>人民币吊牌金额</span><strong>¥576.1m</strong><small>期末库存口径</small></article>
              <article><span>最新OIH</span><strong>已下单未到货</strong><small>按预计到货月份滚动</small></article>
              <article><span>重点风险</span><strong className="red">到货集中</strong><small>8–10月库存与付款承压</small></article>
            </div>
            <section className="panel inventory-table">
              <div className="panel-title"><div><h3>国家库存热力</h3></div><span className="unit">人民币吊牌</span></div>
              {[
                ["示例市场A", 82, "中"],
                ["示例市场B", 76, "高"],
                ["示例市场C", 63, "中"],
                ["示例市场D", 51, "高"],
              ].map(([country, value, risk]) => (
                <div className="country-row" key={country}>
                  <strong>{country}</strong><div><i style={{ width: `${Number(value) / 1.2}%` }} /></div><span>¥{value}m</span><em className={risk === "高" ? "high" : ""}>{risk}风险</em>
                </div>
              ))}
            </section>
          </section>
        )}

        {activeTab === "rules" && (
          <section className="alternate">
            <div className="page-heading"><div><p className="eyebrow blue">OPEN RULEBOOK</p><h2>规则与口径</h2><p>业务和财务使用同一套规则；任何调整必须记录生效日期与审批人。</p></div></div>
            <div className="rules-grid">
              {[
                ["客户锁单", "核验客户订单、不可取消条款、信用额度与回款日期；不直接套用库存周数红线。", "证据不完整 → 有条件通过"],
                ["常规补货", "使用最近13周销量；下单后库存周数超过20周触发黄色、超过26周触发红色。", "WOI > 26周 → 建议调整"],
                ["新品首单", "使用门店数×首铺深度、上市波段和同系列历史表现；新品销量为零不直接否决。", "缺少铺货计划 → 建议调整"],
                ["电商补货", "使用最近4–8周平台销量、活动计划及平台库存；活动订单必须对应营销日历。", "活动证据缺失 → 暂缓"],
                ["国家仓备货", "使用国家总库存、最新OIH、预测准确率和跨国调拨能力；本次订单按新增量模拟。", "批准后库存过高 → 红色预警"],
                ["现金规则", "订单必须落到付款月份；单月超过预算10%或现金安全线时升级审批。", "突破安全线 → 财务否决"],
              ].map(([title, body, rule]) => (
                <article className="rule-card" key={title}><span>{title}</span><p>{body}</p><strong>{rule}</strong></article>
              ))}
            </div>
          </section>
        )}
      </section>
    </main>
  );
}
