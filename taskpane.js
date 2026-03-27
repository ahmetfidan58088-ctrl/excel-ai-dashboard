let cachedData = null;
let cachedAnalysis = "";
let cachedMetrics = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("readBtn").addEventListener("click", readDataFromSheet);
    document.getElementById("analyzeBtn").addEventListener("click", analyzeData);
    document.getElementById("dashboardBtn").addEventListener("click", buildDashboards);
    document.getElementById("writeBtn").addEventListener("click", writeAnalysisToSheet);
    setStatus("Excel hazır.");
  }
});

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function setResult(message) {
  document.getElementById("result").textContent = message;
}

function setBadges(metrics) {
  const el = document.getElementById("summaryBadges");
  const badge = (text, cls) => `<span class="pill ${cls}" style="margin-right:6px;margin-bottom:6px;">${text}</span>`;
  if (!metrics) {
    el.innerHTML = badge("Veri okunmadı", "warn");
    return;
  }
  const items = [];
  items.push(badge(`${metrics.rows} satır`, "good"));
  items.push(badge(`${metrics.columns} kolon`, "good"));
  items.push(badge(`${metrics.blankCellCount} boş hücre`, metrics.blankCellCount > 0 ? "warn" : "good"));
  items.push(badge(`${metrics.issues.length} kritik hata`, metrics.issues.length > 0 ? "bad" : "good"));
  items.push(badge(`${metrics.warnings.length} uyarı`, metrics.warnings.length > 0 ? "warn" : "good"));
  el.innerHTML = items.join("");
}

function normalizeHeader(header) {
  return String(header || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/ı/g, "i")
    .replace(/ğ/g, "g")
    .replace(/ü/g, "u")
    .replace(/ş/g, "s")
    .replace(/ö/g, "o")
    .replace(/ç/g, "c");
}

function findHeaderIndex(headers, possibleNames) {
  const normalizedHeaders = headers.map(normalizeHeader);
  const normalizedPossible = possibleNames.map(normalizeHeader);
  return normalizedHeaders.findIndex((h) => normalizedPossible.includes(h));
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const cleaned = String(value)
    .trim()
    .replace(/\./g, "")
    .replace(/,/g, ".")
    .replace(/[^\d.-]/g, "");
  if (cleaned === "" || cleaned === "-" || cleaned === ".") return null;
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function countBlankCells(rows) {
  let count = 0;
  for (const row of rows) {
    for (const cell of row) {
      if (cell === null || cell === undefined || String(cell).trim() === "") count++;
    }
  }
  return count;
}

async function readDataFromSheet() {
  try {
    setStatus("Veri okunuyor...");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "rowCount", "columnCount", "address"]);
      sheet.load("name");
      await context.sync();

      const values = usedRange.values;
      if (!values || values.length < 2) throw new Error("Aktif sayfada analiz için yeterli veri yok.");

      const headers = values[0].map((h) => String(h).trim());
      const rows = values.slice(1);
      cachedData = {
        sheetName: sheet.name,
        address: usedRange.address,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        headers,
        rows
      };
      cachedMetrics = null;
      cachedAnalysis = "";
      setResult([
        `Sayfa: ${sheet.name}`,
        `Aralık: ${usedRange.address}`,
        `Satır: ${usedRange.rowCount}`,
        `Kolon: ${usedRange.columnCount}`,
        `Başlıklar: ${headers.join(", ")}`
      ].join("\n"));
      setStatus("Veri başarıyla okundu.");
      setBadges(null);
    });
  } catch (error) {
    setStatus("Hata oluştu.");
    setResult(`Hata: ${error.message}`);
  }
}

function analyzeRawData(data, userNotes) {
  const { headers, rows, sheetName, rowCount, columnCount } = data;

  const idx = {
    date: findHeaderIndex(headers, ["tarih", "date", "ay", "month", "donem", "dönem"]),
    dealer: findHeaderIndex(headers, ["bayi", "dealer", "musteri", "müşteri", "account", "kanal"]),
    product: findHeaderIndex(headers, ["urun", "ürün", "product", "model", "sku"]),
    pn: findHeaderIndex(headers, ["pn", "parcano", "partno", "partnumber"]),
    ean: findHeaderIndex(headers, ["ean", "barcode", "barkod"]),
    sales: findHeaderIndex(headers, ["satis", "satış", "sales", "ciro", "tutar", "revenue"]),
    stock: findHeaderIndex(headers, ["stok", "stock", "stokadet", "stokmiktari", "adetstok"]),
    collection: findHeaderIndex(headers, ["tahsilat", "collection", "odeme", "ödeme", "payment"]),
    budget: findHeaderIndex(headers, ["butce", "bütçe", "budget"]),
    actual: findHeaderIndex(headers, ["gerceklesen", "gerçekleşen", "actual"]),
    expense: findHeaderIndex(headers, ["gider", "expense", "masraf", "maliyet"]),
    cash: findHeaderIndex(headers, ["nakit", "cash", "banka", "bank"])
  };

  const issues = [];
  const warnings = [];
  const insights = [];
  const actions = [];
  const blankCellCount = countBlankCells(rows);

  if (blankCellCount > 0) warnings.push(`Boş hücre sayısı: ${blankCellCount}`);

  const duplicateHeaders = headers.filter((h, i) => headers.indexOf(h) !== i);
  if (duplicateHeaders.length > 0) issues.push(`Tekrarlı kolon başlıkları: ${[...new Set(duplicateHeaders)].join(", ")}`);

  const productMap = new Map();
  const dealerMap = new Map();
  const monthMap = new Map();
  const categoryMap = new Map();

  let totalSales = 0;
  let totalStock = 0;
  let totalCollection = 0;
  let totalBudget = 0;
  let totalActual = 0;
  let totalExpense = 0;
  let totalCash = 0;

  let negativeStockCount = 0;
  let salesWithNoStockCount = 0;
  let lowCollectionCount = 0;
  let missingPNEANCount = 0;
  let budgetDeviationCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const dealer = idx.dealer >= 0 ? String(row[idx.dealer] || `Satır ${i + 2}`) : `Satır ${i + 2}`;
    const product = idx.product >= 0 ? String(row[idx.product] || `Ürün yok`) : `Ürün yok`;
    const dateVal = idx.date >= 0 ? String(row[idx.date] || "Belirsiz") : "Belirsiz";
    const monthKey = normalizeMonth(dateVal);
    const sales = idx.sales >= 0 ? toNumber(row[idx.sales]) : null;
    const stock = idx.stock >= 0 ? toNumber(row[idx.stock]) : null;
    const collection = idx.collection >= 0 ? toNumber(row[idx.collection]) : null;
    const budget = idx.budget >= 0 ? toNumber(row[idx.budget]) : null;
    const actual = idx.actual >= 0 ? toNumber(row[idx.actual]) : null;
    const expense = idx.expense >= 0 ? toNumber(row[idx.expense]) : null;
    const cash = idx.cash >= 0 ? toNumber(row[idx.cash]) : null;
    const pn = idx.pn >= 0 ? String(row[idx.pn] || "").trim() : "";
    const ean = idx.ean >= 0 ? String(row[idx.ean] || "").trim() : "";

    if (!dealerMap.has(dealer)) dealerMap.set(dealer, { dealer, sales: 0, stock: 0, collection: 0, rows: 0 });
    if (!productMap.has(product)) productMap.set(product, { product, sales: 0, stock: 0, rows: 0 });
    if (!monthMap.has(monthKey)) monthMap.set(monthKey, { month: monthKey, sales: 0, expense: 0, actual: 0, budget: 0 });

    const d = dealerMap.get(dealer);
    const p = productMap.get(product);
    const m = monthMap.get(monthKey);

    d.rows += 1;
    p.rows += 1;

    if (sales !== null) {
      totalSales += sales;
      d.sales += sales;
      p.sales += sales;
      m.sales += sales;
    }
    if (stock !== null) {
      totalStock += stock;
      d.stock += stock;
      p.stock += stock;
    }
    if (collection !== null) {
      totalCollection += collection;
      d.collection += collection;
    }
    if (budget !== null) {
      totalBudget += budget;
      m.budget += budget;
    }
    if (actual !== null) {
      totalActual += actual;
      m.actual += actual;
    }
    if (expense !== null) {
      totalExpense += expense;
      m.expense += expense;
    }
    if (cash !== null) totalCash += cash;

    if (stock !== null && stock < 0) negativeStockCount++;
    if (sales !== null && sales > 0 && stock !== null && stock <= 0) salesWithNoStockCount++;
    if (sales !== null && sales > 0 && collection !== null && collection < sales * 0.5) lowCollectionCount++;
    if ((idx.pn >= 0 || idx.ean >= 0) && (!pn || !ean)) missingPNEANCount++;
    if (budget !== null && budget !== 0 && actual !== null && Math.abs((actual - budget) / budget) > 0.2) budgetDeviationCount++;
  }

  if (idx.sales === -1) warnings.push("Satış/Ciro kolonu bulunamadı.");
  if (idx.stock === -1) warnings.push("Stok kolonu bulunamadı.");
  if (idx.dealer === -1) warnings.push("Bayi/Müşteri kolonu bulunamadı.");
  if (idx.product === -1) warnings.push("Ürün/Model kolonu bulunamadı.");
  if (idx.collection === -1) warnings.push("Tahsilat kolonu bulunamadı.");
  if (idx.budget === -1 || idx.actual === -1) warnings.push("Bütçe-Gerçekleşen karşılaştırması için kolonlar eksik.");

  if (negativeStockCount > 0) issues.push(`Negatif stok satır sayısı: ${negativeStockCount}`);
  if (salesWithNoStockCount > 0) issues.push(`Satış var ama stok 0/negatif satır sayısı: ${salesWithNoStockCount}`);
  if (lowCollectionCount > 0) warnings.push(`Tahsilatı zayıf satır sayısı: ${lowCollectionCount}`);
  if (missingPNEANCount > 0) warnings.push(`PN/EAN eksiği olan satır sayısı: ${missingPNEANCount}`);
  if (budgetDeviationCount > 0) warnings.push(`Bütçeden %20+ sapan satır sayısı: ${budgetDeviationCount}`);

  const dealers = Array.from(dealerMap.values()).sort((a, b) => b.sales - a.sales);
  const products = Array.from(productMap.values()).sort((a, b) => b.sales - a.sales);
  const months = Array.from(monthMap.values()).sort((a, b) => a.month.localeCompare(b.month));

  const topDealers = dealers.slice(0, 10);
  const weakDealers = dealers.filter((x) => x.stock > 0 && x.sales <= 0).slice(0, 10);
  const topProducts = products.slice(0, 10);
  const riskyProducts = products.filter((x) => x.stock > 0 && x.sales <= 0).slice(0, 10);

  insights.push(`Analiz edilen sayfa: ${sheetName}`);
  insights.push(`Toplam veri boyutu: ${rowCount} satır x ${columnCount} kolon`);
  insights.push(`Toplam ciro: ${formatNumber(totalSales)}`);
  insights.push(`Toplam stok: ${formatNumber(totalStock)}`);
  insights.push(`Toplam tahsilat: ${formatNumber(totalCollection)}`);
  insights.push(`Toplam gider: ${formatNumber(totalExpense)}`);
  insights.push(`Net fark: ${formatNumber(totalSales - totalExpense)}`);

  if (topDealers.length) insights.push(`En yüksek satışlı ilk bayi: ${topDealers[0].dealer} (${formatNumber(topDealers[0].sales)})`);
  if (topProducts.length) insights.push(`En yüksek satışlı ilk ürün: ${topProducts[0].product} (${formatNumber(topProducts[0].sales)})`);
  if (!issues.length) insights.push("Kritik veri hatası tespit edilmedi.");

  actions.push("Eksik/boş hücreler kontrol edilmeli.");
  if (negativeStockCount > 0) actions.push("Negatif stok görünen satırlar doğrulanmalı.");
  if (salesWithNoStockCount > 0) actions.push("Satış olup stok görünmeyen satırlar incelenmeli.");
  if (lowCollectionCount > 0) actions.push("Tahsilatı düşük bayi kayıtları ayrı raporlanmalı.");
  if (missingPNEANCount > 0) actions.push("PN/EAN eksikleri için düzeltme listesi oluşturulmalı.");
  if (userNotes && userNotes.trim()) actions.push(`Kullanıcı notu dikkate alınmalı: ${userNotes.trim()}`);

  const metrics = {
    sheetName,
    rows: rows.length,
    columns: headers.length,
    blankCellCount,
    issues,
    warnings,
    insights,
    actions,
    totals: {
      sales: totalSales,
      stock: totalStock,
      collection: totalCollection,
      budget: totalBudget,
      actual: totalActual,
      expense: totalExpense,
      cash: totalCash,
      net: totalSales - totalExpense,
      collectionRate: totalSales > 0 ? totalCollection / totalSales : 0,
      budgetRate: totalBudget > 0 ? totalActual / totalBudget : 0
    },
    indices: idx,
    headers,
    topDealers,
    weakDealers,
    topProducts,
    riskyProducts,
    monthSeries: months,
    dealerSeries: dealers,
    productSeries: products
  };

  const report = [
    "AI DASHBOARD ONLINE V4 - ÖN ANALİZ",
    "===================================",
    "",
    "1. GENEL ÖZET",
    ...insights.map((x) => `- ${x}`),
    "",
    "2. KRİTİK HATALAR",
    ...(issues.length ? issues.map((x) => `- ${x}`) : ["- Kritik hata bulunmadı."]),
    "",
    "3. UYARILAR",
    ...(warnings.length ? warnings.map((x) => `- ${x}`) : ["- Uyarı bulunmadı."]),
    "",
    "4. AKSİYONLAR",
    ...actions.map((x) => `- ${x}`)
  ].join("\n");

  return { report, metrics };
}

function formatNumber(value) {
  const n = Number(value || 0);
  return n.toLocaleString("tr-TR", { maximumFractionDigits: 2 });
}

function normalizeMonth(input) {
  const raw = String(input || "Belirsiz").trim();
  if (!raw) return "Belirsiz";
  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) {
    const month = String(d.getMonth() + 1).padStart(2, "0");
    return `${d.getFullYear()}-${month}`;
  }
  const m = raw.match(/(20\d{2}).{0,3}(\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}`;
  return raw.slice(0, 20);
}

async function analyzeData() {
  try {
    if (!cachedData) {
      setStatus("Önce veri okunmalı.");
      setResult("Önce 'Veriyi Oku' butonuna bas.");
      return;
    }
    setStatus("Analiz yapılıyor...");
    const notes = document.getElementById("notes").value || "";
    const { report, metrics } = analyzeRawData(cachedData, notes);
    cachedMetrics = metrics;
    cachedAnalysis = report;
    setResult(report);
    setBadges(metrics);
    setStatus("Analiz tamamlandı.");
  } catch (error) {
    setStatus("Analiz sırasında hata oluştu.");
    setResult(`Hata: ${error.message}`);
  }
}

async function writeAnalysisToSheet() {
  try {
    if (!cachedAnalysis) {
      setStatus("Önce analiz yapmalısın.");
      return;
    }
    setStatus("AI_ANALYSIS yazılıyor...");
    await Excel.run(async (context) => {
      const sheet = await getOrCreateSheet(context, "AI_ANALYSIS", true);
      const lines = cachedAnalysis.split("\n").map((line) => [line]);
      const range = sheet.getRange(`A1:A${lines.length}`);
      range.values = lines;
      range.format.autofitColumns();
      range.getCell(0, 0).format.font.bold = true;
      range.getCell(0, 0).format.font.size = 14;
      await context.sync();
    });
    setStatus("AI_ANALYSIS yazıldı.");
  } catch (error) {
    setStatus("Yazma hatası.");
    setResult(`Hata: ${error.message}`);
  }
}

async function buildDashboards() {
  try {
    if (!cachedMetrics) {
      setStatus("Önce analiz yapmalısın.");
      setResult("Önce 'Analiz Yap' butonuna bas.");
      return;
    }
    setStatus("Dashboard sayfaları kuruluyor...");
    await Excel.run(async (context) => {
      await createDashboardDataSheet(context, cachedMetrics);
      await createExecutiveDashboard(context, cachedMetrics);
      await createSalesDashboard(context, cachedMetrics);
      await createStockDashboard(context, cachedMetrics);
      await createFinanceDashboard(context, cachedMetrics);
      await createDealerDashboard(context, cachedMetrics);
      await createAnalystScreen(context, cachedMetrics);
      await context.sync();
    });
    setStatus("Dashboardlar oluşturuldu.");
    setResult(`${cachedAnalysis}\n\nDashboard sayfaları başarıyla oluşturuldu.`);
  } catch (error) {
    setStatus("Dashboard kurulumunda hata oluştu.");
    setResult(`Hata: ${error.message}`);
  }
}

async function getOrCreateSheet(context, name, clear = false) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  let sheet = sheets.items.find((s) => s.name === name);
  if (!sheet) sheet = sheets.add(name);
  if (clear) {
    const used = sheet.getUsedRangeOrNullObject();
    used.load("address");
    await context.sync();
    if (!used.isNullObject) used.clear();
  }
  return sheet;
}

function styleHeader(range, bg = "#1f4fd1", font = "#ffffff") {
  range.format.fill.color = bg;
  range.format.font.color = font;
  range.format.font.bold = true;
}

function setTitle(sheet, title, subtitle) {
  const titleRange = sheet.getRange("A1:H1");
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format.font.bold = true;
  titleRange.format.font.size = 18;
  titleRange.format.fill.color = "#163caa";
  titleRange.format.font.color = "#ffffff";
  const sub = sheet.getRange("A2:H2");
  sub.merge();
  sub.values = [[subtitle || ""]];
  sub.format.fill.color = "#eaf0ff";
  sub.format.font.color = "#334155";
}

function writeKpiCard(sheet, cell, label, value, color) {
  const labelRange = sheet.getRange(cell);
  labelRange.values = [[label]];
  labelRange.format.fill.color = color;
  labelRange.format.font.color = "#ffffff";
  labelRange.format.font.bold = true;

  const [col, row] = splitCell(cell);
  const valueCell = `${col}${row + 1}`;
  const vRange = sheet.getRange(valueCell);
  vRange.values = [[value]];
  vRange.format.font.bold = true;
  vRange.format.font.size = 14;
  vRange.format.fill.color = "#f8fafc";
  vRange.format.borders.getItem("EdgeBottom").style = "Continuous";
}

function splitCell(cell) {
  const m = cell.match(/^([A-Z]+)(\d+)$/);
  return [m[1], Number(m[2])];
}

async function createDashboardDataSheet(context, metrics) {
  const sheet = await getOrCreateSheet(context, "Dashboard_Data", true);
  setTitle(sheet, "Dashboard_Data", "Özet veri katmanı");

  const summary = [
    ["Metric", "Value"],
    ["Total Sales", metrics.totals.sales],
    ["Total Stock", metrics.totals.stock],
    ["Total Collection", metrics.totals.collection],
    ["Total Expense", metrics.totals.expense],
    ["Net", metrics.totals.net],
    ["Collection Rate", metrics.totals.collectionRate],
    ["Budget Rate", metrics.totals.budgetRate],
    ["Rows", metrics.rows],
    ["Columns", metrics.columns],
    ["Blank Cells", metrics.blankCellCount],
    ["Issues", metrics.issues.length],
    ["Warnings", metrics.warnings.length]
  ];
  const r1 = sheet.getRangeByIndexes(3, 0, summary.length, 2);
  r1.values = summary;
  styleHeader(r1.getRow(0));

  const dealer = [["Dealer", "Sales", "Stock", "Collection"]].concat(
    metrics.topDealers.slice(0, 10).map((x) => [x.dealer, x.sales, x.stock, x.collection])
  );
  const r2 = sheet.getRangeByIndexes(3, 4, dealer.length, 4);
  r2.values = dealer;
  styleHeader(r2.getRow(0));

  const product = [["Product", "Sales", "Stock"]].concat(
    metrics.topProducts.slice(0, 10).map((x) => [x.product, x.sales, x.stock])
  );
  const r3 = sheet.getRangeByIndexes(16, 0, product.length, 3);
  r3.values = product;
  styleHeader(r3.getRow(0));

  const month = [["Month", "Sales", "Expense", "Actual", "Budget"]].concat(
    metrics.monthSeries.slice(0, 24).map((x) => [x.month, x.sales, x.expense, x.actual, x.budget])
  );
  const r4 = sheet.getRangeByIndexes(16, 4, month.length, 5);
  r4.values = month;
  styleHeader(r4.getRow(0));

  sheet.getUsedRange().format.autofitColumns();
}

async function createExecutiveDashboard(context, metrics) {
  const sheet = await getOrCreateSheet(context, "EXECUTIVE_DASHBOARD", true);
  setTitle(sheet, "EXECUTIVE_DASHBOARD", "Yönetici özeti");
  writeKpiCard(sheet, "A4", "Toplam Ciro", formatNumber(metrics.totals.sales), "#2458e7");
  writeKpiCard(sheet, "C4", "Toplam Gider", formatNumber(metrics.totals.expense), "#ef4444");
  writeKpiCard(sheet, "E4", "Net", formatNumber(metrics.totals.net), "#0f9d58");
  writeKpiCard(sheet, "G4", "Tahsilat Oranı", `${(metrics.totals.collectionRate * 100).toFixed(1)}%`, "#7c3aed");

  const summary = [
    ["Alan", "Değer"],
    ["Toplam Stok", metrics.totals.stock],
    ["Toplam Tahsilat", metrics.totals.collection],
    ["Bütçe Gerçekleşme", `${(metrics.totals.budgetRate * 100).toFixed(1)}%`],
    ["Kritik Hata", metrics.issues.length],
    ["Uyarı", metrics.warnings.length],
    ["Aktif Bayi", metrics.dealerSeries.length]
  ];
  const table = sheet.getRangeByIndexes(8, 0, summary.length, 2);
  table.values = summary;
  styleHeader(table.getRow(0));

  const warnData = [["Kritik Uyarılar"]]
    .concat((metrics.issues.length ? metrics.issues : ["Kritik hata bulunmadı."]).map((x) => [x]))
    .concat([["Uyarılar"]])
    .concat((metrics.warnings.length ? metrics.warnings : ["Uyarı bulunmadı."]).slice(0, 8).map((x) => [x]));
  const warnRange = sheet.getRangeByIndexes(8, 3, warnData.length, 4);
  warnRange.values = warnData.map((r) => [r[0], "", "", ""]);
  warnRange.getRow(0).format.fill.color = "#fff4d6";
  warnRange.getRow(0).format.font.bold = true;
  warnRange.getCell(metrics.issues.length + 1, 0).format.fill.color = "#eef2ff";
  warnRange.getCell(metrics.issues.length + 1, 0).format.font.bold = true;

  sheet.getUsedRange().format.autofitColumns();
}

async function createSalesDashboard(context, metrics) {
  const sheet = await getOrCreateSheet(context, "SALES_DASHBOARD", true);
  setTitle(sheet, "SALES_DASHBOARD", "Satış görünümü");
  writeKpiCard(sheet, "A4", "Toplam Satış", formatNumber(metrics.totals.sales), "#2458e7");
  writeKpiCard(sheet, "C4", "Ürün Sayısı", metrics.productSeries.length, "#0f9d58");
  writeKpiCard(sheet, "E4", "İlk Bayi Satışı", metrics.topDealers[0] ? formatNumber(metrics.topDealers[0].sales) : 0, "#7c3aed");
  writeKpiCard(sheet, "G4", "İlk Ürün Satışı", metrics.topProducts[0] ? formatNumber(metrics.topProducts[0].sales) : 0, "#d97706");

  const topProducts = [["Top 10 Ürün", "Satış", "Stok"]].concat(metrics.topProducts.slice(0, 10).map((x) => [x.product, x.sales, x.stock]));
  const rp = sheet.getRangeByIndexes(8, 0, topProducts.length, 3);
  rp.values = topProducts;
  styleHeader(rp.getRow(0));

  const topDealers = [["Top 10 Bayi", "Satış", "Stok", "Tahsilat"]].concat(metrics.topDealers.slice(0, 10).map((x) => [x.dealer, x.sales, x.stock, x.collection]));
  const rd = sheet.getRangeByIndexes(8, 4, topDealers.length, 4);
  rd.values = topDealers;
  styleHeader(rd.getRow(0));

  const months = [["Ay", "Satış"]].concat(metrics.monthSeries.slice(0, 24).map((x) => [x.month, x.sales]));
  const rm = sheet.getRangeByIndexes(21, 0, months.length, 2);
  rm.values = months;
  styleHeader(rm.getRow(0));
  sheet.getUsedRange().format.autofitColumns();
}

async function createStockDashboard(context, metrics) {
  const sheet = await getOrCreateSheet(context, "STOCK_DASHBOARD", true);
  setTitle(sheet, "STOCK_DASHBOARD", "Stok görünümü");
  writeKpiCard(sheet, "A4", "Toplam Stok", formatNumber(metrics.totals.stock), "#2458e7");
  writeKpiCard(sheet, "C4", "Riskli Ürün", metrics.riskyProducts.length, "#ef4444");
  writeKpiCard(sheet, "E4", "Negatif Stok Hata", metrics.issues.filter(x => x.toLowerCase().includes("negatif stok")).length, "#d97706");
  writeKpiCard(sheet, "G4", "Stok/Satış Çelişkisi", metrics.issues.filter(x => x.toLowerCase().includes("stok 0")).length, "#7c3aed");

  const risky = [["Riskli Ürün", "Satış", "Stok"]].concat(metrics.riskyProducts.slice(0, 10).map((x) => [x.product, x.sales, x.stock]));
  const rr = sheet.getRangeByIndexes(8, 0, risky.length, 3);
  rr.values = risky;
  styleHeader(rr.getRow(0), "#b91c1c");

  const top = [["En Yüksek Stoklu Ürün", "Satış", "Stok"]].concat(
    [...metrics.productSeries].sort((a, b) => b.stock - a.stock).slice(0, 10).map((x) => [x.product, x.sales, x.stock])
  );
  const rt = sheet.getRangeByIndexes(8, 4, top.length, 3);
  rt.values = top;
  styleHeader(rt.getRow(0), "#0f766e");
  sheet.getUsedRange().format.autofitColumns();
}

async function createFinanceDashboard(context, metrics) {
  const sheet = await getOrCreateSheet(context, "FINANCE_DASHBOARD", true);
  setTitle(sheet, "FINANCE_DASHBOARD", "Finans görünümü");
  writeKpiCard(sheet, "A4", "Toplam Tahsilat", formatNumber(metrics.totals.collection), "#2458e7");
  writeKpiCard(sheet, "C4", "Toplam Gider", formatNumber(metrics.totals.expense), "#ef4444");
  writeKpiCard(sheet, "E4", "Net", formatNumber(metrics.totals.net), "#0f9d58");
  writeKpiCard(sheet, "G4", "Bütçe Gerçekleşme", `${(metrics.totals.budgetRate * 100).toFixed(1)}%`, "#7c3aed");

  const months = [["Ay", "Satış", "Gider", "Net", "Gerçekleşen", "Bütçe"]].concat(
    metrics.monthSeries.slice(0, 24).map((x) => [x.month, x.sales, x.expense, x.sales - x.expense, x.actual, x.budget])
  );
  const range = sheet.getRangeByIndexes(8, 0, months.length, 6);
  range.values = months;
  styleHeader(range.getRow(0));
  sheet.getUsedRange().format.autofitColumns();
}

async function createDealerDashboard(context, metrics) {
  const sheet = await getOrCreateSheet(context, "DEALER_DASHBOARD", true);
  setTitle(sheet, "DEALER_DASHBOARD", "Bayi görünümü");
  writeKpiCard(sheet, "A4", "Aktif Bayi", metrics.dealerSeries.length, "#2458e7");
  writeKpiCard(sheet, "C4", "İlk Bayi Satışı", metrics.topDealers[0] ? formatNumber(metrics.topDealers[0].sales) : 0, "#0f9d58");
  writeKpiCard(sheet, "E4", "Zayıf Bayi", metrics.weakDealers.length, "#ef4444");
  writeKpiCard(sheet, "G4", "Tahsilat Oranı", `${(metrics.totals.collectionRate * 100).toFixed(1)}%`, "#7c3aed");

  const top = [["Top Bayi", "Satış", "Stok", "Tahsilat"]].concat(metrics.topDealers.slice(0, 10).map((x) => [x.dealer, x.sales, x.stock, x.collection]));
  const rt = sheet.getRangeByIndexes(8, 0, top.length, 4);
  rt.values = top;
  styleHeader(rt.getRow(0));

  const weak = [["Sorunlu Bayi", "Satış", "Stok", "Tahsilat"]].concat(metrics.weakDealers.slice(0, 10).map((x) => [x.dealer, x.sales, x.stock, x.collection]));
  const rw = sheet.getRangeByIndexes(8, 5, Math.max(weak.length, 2), 4);
  rw.values = weak.length > 1 ? weak : [["Sorunlu Bayi", "Satış", "Stok", "Tahsilat"], ["Yok", 0, 0, 0]];
  styleHeader(rw.getRow(0), "#b91c1c");
  sheet.getUsedRange().format.autofitColumns();
}

async function createAnalystScreen(context, metrics) {
  const sheet = await getOrCreateSheet(context, "ANALYST_SCREEN", true);
  setTitle(sheet, "ANALYST_SCREEN", "Analist kontrol ekranı");
  writeKpiCard(sheet, "A4", "Boş Hücre", metrics.blankCellCount, "#d97706");
  writeKpiCard(sheet, "C4", "Kritik Hata", metrics.issues.length, "#ef4444");
  writeKpiCard(sheet, "E4", "Uyarı", metrics.warnings.length, "#2458e7");
  writeKpiCard(sheet, "G4", "Mail/İş Listesi", metrics.actions.length, "#7c3aed");

  const issues = [["Veri Hataları"]].concat((metrics.issues.length ? metrics.issues : ["Kritik hata bulunmadı."]).map((x) => [x]));
  const ri = sheet.getRangeByIndexes(8, 0, issues.length, 3);
  ri.values = issues.map((r) => [r[0], "", ""]);
  ri.getRow(0).format.fill.color = "#fee2e2";
  ri.getRow(0).format.font.bold = true;

  const warnings = [["Uyarılar"]].concat((metrics.warnings.length ? metrics.warnings : ["Uyarı bulunmadı."]).map((x) => [x]));
  const rw = sheet.getRangeByIndexes(8, 4, warnings.length, 4);
  rw.values = warnings.map((r) => [r[0], "", "", ""]);
  rw.getRow(0).format.fill.color = "#fef3c7";
  rw.getRow(0).format.font.bold = true;

  const actions = [["Operasyon Aksiyonları"]].concat(metrics.actions.map((x) => [x]));
  const ra = sheet.getRangeByIndexes(20, 0, actions.length, 6);
  ra.values = actions.map((r) => [r[0], "", "", "", "", ""]);
  ra.getRow(0).format.fill.color = "#dbeafe";
  ra.getRow(0).format.font.bold = true;

  const matchRows = [
    ["Kontrol Alanı", "Durum"],
    ["PN kolonu", metrics.indices.pn >= 0 ? "Var" : "Yok"],
    ["EAN kolonu", metrics.indices.ean >= 0 ? "Var" : "Yok"],
    ["Bayi kolonu", metrics.indices.dealer >= 0 ? "Var" : "Yok"],
    ["Ürün kolonu", metrics.indices.product >= 0 ? "Var" : "Yok"],
    ["Satış kolonu", metrics.indices.sales >= 0 ? "Var" : "Yok"],
    ["Stok kolonu", metrics.indices.stock >= 0 ? "Var" : "Yok"]
  ];
  const rm = sheet.getRangeByIndexes(20, 7, matchRows.length, 2);
  rm.values = matchRows;
  styleHeader(rm.getRow(0), "#0f766e");
  sheet.getUsedRange().format.autofitColumns();
}
