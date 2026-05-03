@@ -31,79 +31,81 @@ async function analyzeAllSheets() {
                    allData.push({ sheetName: sheet.name, headers, rows });
                }
            }

            if (allData.length === 0) {
                throw new Error("Hiç veri bulunamadı.");
            }

            const columnMapping = detectColumnsAcrossSheets(allData);
            const mergedData = { rows: [], headers: [] };
            for (const data of allData) {
                const mapped = mapDataToColumns(data.rows, data.headers, columnMapping);
                mergedData.rows.push(...mapped.rows);
                if (mergedData.headers.length === 0 && mapped.headers.length) {
                    mergedData.headers = mapped.headers;
                }
            }

            const qualityIssues = runQualityChecks(mergedData, columnMapping);
            await createDashboardSheets(context, mergedData, columnMapping, qualityIssues);
            
            const resultText = `✅ Analiz tamamlandı!\n\n` +
                `📊 Toplam ${mergedData.rows.length} satır veri işlendi.\n` +
                `🔍 Tespit edilen kolonlar: ${Object.entries(columnMapping).map(([k,v]) => `${k}: ${v || "bulunamadı"}`).join(", ")}\n` +
                `⚠️ ${qualityIssues.length} adet veri kalite sorunu tespit edildi.\n\n` +
                `📌 Dashboard sayfaları oluşturuldu: 00_Executive, 01_Sales, 02_Stock, 03_Finance, 04_Channel, 05_Product, 06_DataQuality`;
                `📌 Dashboard sayfaları oluşturuldu: 00_Executive, 01_Sales, 02_Stock, 03_Finance, 04_Channel, 05_Product, 06_DataQuality, 07_MonthlyTrend`;
            showResult(resultText);

            await context.sync();
        });
    } catch (error) {
        console.error("Hata:", error);
        showError("Analiz sırasında hata: " + error.message);
    } finally {
        showLoading(false);
    }
}

// ========== KOLON TANIMA (Fuzzy + Tip Analizi) ==========
const ALIASES = {
    date: ["tarih", "date", "islem_tarihi", "siparis_tarihi", "invoice_date", "month", "ay"],
    product: ["urun", "ürün", "product", "model", "malzeme", "item", "sku", "pn", "part_number"],
    quantity: ["adet", "miktar", "quantity", "qty", "satilan_adet", "satis_adedi", "units"],
    revenue: ["ciro", "revenue", "sales_amount", "tutar", "net_satis", "net_satis"],
    stock: ["stok", "stock", "inventory", "mevcut_stok"],
    budget: ["butce", "bütçe", "budget", "plan"],
    actual: ["gerceklesen", "gerçekleşen", "actual", "realized"],
    cost: ["maliyet", "cost", "gider", "expense"],
    channel: ["kanal", "bayi", "channel", "dealer", "customer", "müşteri"],
    region: ["bolge", "bölge", "region"],
    status: ["durum", "status", "state"],
    phase: ["faz", "phase", "asama"],
    projectType: ["proje tipi", "project type", "tip"],
    safetyIncidents: ["güvenlik", "safety", "incident", "olay"]
    safetyIncidents: ["güvenlik", "safety", "incident", "olay"],
    sku: ["sku", "stok kodu", "urun kodu", "ürün kodu", "part no"],
    ean: ["ean", "barkod", "barcode", "gtin"]
};

function normalizeString(s) {
    if (!s) return "";
    return s.toLowerCase()
        .replace(/ç/g, "c").replace(/ğ/g, "g").replace(/ı/g, "i").replace(/ö/g, "o").replace(/ş/g, "s").replace(/ü/g, "u")
        .replace(/[^a-z0-9]/g, " ")
        .trim();
}

function similarityScore(str1, str2) {
    const tokens1 = str1.split(/\s+/);
    const tokens2 = str2.split(/\s+/);
    let match = 0;
    for (let t of tokens1) {
        if (tokens2.includes(t)) match++;
    }
    return match / Math.max(tokens1.length, tokens2.length);
}

function detectColumnsAcrossSheets(allData) {
    const mapping = {};
    for (let [canonical, aliases] of Object.entries(ALIASES)) {
        mapping[canonical] = null;
    }
@@ -175,71 +177,74 @@ function parseDate(val) {
    let day, month, year;
    if (str.includes(".")) {
        [day, month, year] = str.split(".");
    } else if (str.includes("-")) {
        [year, month, day] = str.split("-");
    } else {
        return null;
    }
    const d = new Date(year, month-1, day);
    return isNaN(d.getTime()) ? null : d;
}

function parseNumber(val) {
    if (val === undefined || val === null) return NaN;
    if (typeof val === "number") return val;
    let s = String(val).replace(/[^0-9,\.\-]/g, "").replace(",", ".");
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
}

// ========== VERİ KALİTE KONTROLLERİ ==========
function runQualityChecks(mergedData, mapping) {
    const issues = [];
    const rows = mergedData.rows;
    if (rows.length === 0) return issues;
    const mappedColumns = Object.entries(mapping)
        .filter(([, header]) => header)
        .map(([canonical]) => canonical);

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        for (let col of Object.keys(row)) {
        for (let col of mappedColumns) {
            if (row[col] === undefined || row[col] === null || row[col] === "") {
                issues.push({
                    sheet: "Tüm veri",
                    row: i+2,
                    column: col,
                    issue: "Eksik değer",
                    severity: row[col] === null ? "medium" : "low",
                    suggestion: "Hücreyi doldurun veya varsayılan değer atayın."
                });
            }
        }
    }

    if (mapping.date) {
        for (let i = 0; i < rows.length; i++) {
            const d = rows[i].date;
            if (d === null && rows[i].date !== undefined && rows[i].date !== "") {
            if (d === null) {
                issues.push({
                    sheet: "Tüm veri",
                    row: i+2,
                    column: mapping.date,
                    issue: "Geçersiz tarih formatı",
                    severity: "medium",
                    suggestion: "Tarih formatını GG.AA.YYYY veya YYYY-AA-GG olarak düzeltin."
                });
            }
        }
    }

    const numericCols = ["quantity", "revenue", "stock", "budget", "actual", "cost"];
    for (let col of numericCols) {
        if (mapping[col]) {
            for (let i = 0; i < rows.length; i++) {
                const val = rows[i][col];
                if (val !== undefined && val !== null && val !== "" && isNaN(val)) {
                    issues.push({
                        sheet: "Tüm veri",
                        row: i+2,
                        column: mapping[col],
                        issue: "Sayısal olmayan değer",
                        severity: "high",
                        suggestion: "Değeri sayıya çevirin (virgül, TL gibi işaretleri temizleyin)."
@@ -273,74 +278,75 @@ function runQualityChecks(mergedData, mapping) {
            const ean = rows[i].ean;
            if (sku && ean) {
                if (!groups.has(sku)) groups.set(sku, new Set());
                groups.get(sku).add(ean);
            }
        }
        for (let [sku, eans] of groups.entries()) {
            if (eans.size > 1) {
                issues.push({
                    sheet: "Tüm veri",
                    row: -1,
                    column: mapping.sku,
                    issue: "Aynı SKU'ya birden fazla EAN atanmış",
                    severity: "high",
                    suggestion: `SKU ${sku} için EAN'leri birleştirin veya düzeltin.`
                });
            }
        }
    }

    return issues;
}

// ========== DASHBOARD SAYFALARI OLUŞTURMA ==========
async function createDashboardSheets(context, data, mapping, issues) {
    const sheetNames = ["00_Executive", "01_Sales", "02_Stock", "03_Finance", "04_Channel", "05_Product", "06_DataQuality"];
    const sheetNames = ["00_Executive", "01_Sales", "02_Stock", "03_Finance", "04_Channel", "05_Product", "06_DataQuality", "07_MonthlyTrend"];
    
    // Mevcut sayfaları güvenli şekilde sil
    for (let name of sheetNames) {
        try {
            const sheet = context.workbook.worksheets.getItem(name);
            sheet.load("name");
            await context.sync();
            sheet.delete();
            await context.sync();
        } catch (e) {
            // Sayfa yoksa sessizce geç
            console.log(`${name} sayfası zaten yok veya silinemiyor.`);
        }
    }

    // Yeni sayfaları oluştur
    await createExecutiveSheet(context, data, mapping, issues);
    await createSalesSheet(context, data, mapping);
    await createStockSheet(context, data, mapping);
    await createFinanceSheet(context, data, mapping);
    await createChannelSheet(context, data, mapping);
    await createProductSheet(context, data, mapping);
    await createQualitySheet(context, issues);
    await createMonthlyTrendSheet(context, data, mapping);
}

async function createExecutiveSheet(context, data, mapping, issues) {
    const sheet = context.workbook.worksheets.add("00_Executive");
    sheet.getRange("A1").values = [["EXECUTIVE DASHBOARD - ÖZET"]];
    sheet.getRange("A1").format.font.bold = true;
    let row = 2;

    const totalQty = data.rows.reduce((sum, r) => sum + (isNaN(r.quantity) ? 0 : r.quantity), 0);
    const totalRevenue = data.rows.reduce((sum, r) => sum + (isNaN(r.revenue) ? 0 : r.revenue), 0);
    const avgQty = data.rows.length ? totalQty / data.rows.length : 0;

    sheet.getRange(row, 0).values = [["Toplam Adet", totalQty]];
    sheet.getRange(row+1, 0).values = [["Toplam Ciro (TL)", totalRevenue]];
    sheet.getRange(row+2, 0).values = [["Ortalama Adet", avgQty.toFixed(2)]];
    sheet.getRange(row+3, 0).values = [["Kalite Sorunu Sayısı", issues.length]];
    row += 5;

    if (mapping.product && mapping.quantity) {
        const prodMap = new Map();
        for (const r of data.rows) {
            if (r.product && !isNaN(r.quantity)) {
                prodMap.set(r.product, (prodMap.get(r.product) || 0) + r.quantity);
            }
        }
@@ -490,53 +496,97 @@ async function createProductSheet(context, data, mapping) {
        for (let p of products) {
            sheet.getRange(row, 0).values = [[p[0], p[1]]];
            row++;
        }
    } else {
        sheet.getRange(row, 0).values = [["Ürün veya adet sütunu bulunamadı."]];
    }
    sheet.getRange("A:C").format.autofitColumns();
}

async function createQualitySheet(context, issues) {
    const sheet = context.workbook.worksheets.add("06_DataQuality");
    sheet.getRange("A1").values = [["VERİ KALİTE RAPORU"]];
    sheet.getRange("A1").format.font.bold = true;
    let row = 2;

    sheet.getRange(row, 0).values = [["Sayfa", "Satır", "Sütun", "Sorun", "Şiddet", "Öneri"]];
    row++;
    for (let issue of issues) {
        sheet.getRange(row, 0).values = [[issue.sheet, issue.row, issue.column, issue.issue, issue.severity, issue.suggestion]];
        row++;
    }
    sheet.getRange("A:F").format.autofitColumns();
}

async function createMonthlyTrendSheet(context, data, mapping) {
    const sheet = context.workbook.worksheets.add("07_MonthlyTrend");
    sheet.getRange("A1").values = [["AYLIK TREND RAPORU"]];
    sheet.getRange("A1").format.font.bold = true;
    let row = 2;

    if (!(mapping.date && (mapping.revenue || mapping.quantity))) {
        sheet.getRange(row, 0).values = [["Aylık trend için tarih ve ciro/adet sütunu bulunamadı."]];
        sheet.getRange("A:C").format.autofitColumns();
        return;
    }

    const monthlyMap = new Map();
    for (const r of data.rows) {
        if (!(r.date instanceof Date) || isNaN(r.date.getTime())) continue;
        const key = `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`;
        const existing = monthlyMap.get(key) || { revenue: 0, quantity: 0 };
        if (!isNaN(r.revenue)) existing.revenue += r.revenue;
        if (!isNaN(r.quantity)) existing.quantity += r.quantity;
        monthlyMap.set(key, existing);
    }

    const rows = Array.from(monthlyMap.entries())
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([month, values]) => [month, values.quantity, values.revenue]);

    if (rows.length === 0) {
        sheet.getRange(row, 0).values = [["Geçerli tarih içeren veri bulunamadı."]];
        sheet.getRange("A:C").format.autofitColumns();
        return;
    }

    sheet.getRange(row, 0).values = [["Ay", "Toplam Adet", "Toplam Ciro"]];
    row++;
    sheet.getRange(`A${row}:C${row + rows.length - 1}`).values = rows;

    const dataRange = sheet.getRange(`A${row - 1}:C${row + rows.length - 1}`);
    const chart = sheet.charts.add("line", dataRange, "auto");
    chart.title.text = "Aylık Satış Trendi";
    chart.legend.position = "bottom";

    sheet.getRange("A:C").format.autofitColumns();
}

// ========== UI YARDIMCILARI ==========
function showLoading(show) {
    const loading = document.getElementById("loading");
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (loading) loading.classList.toggle("hidden", !show);
    if (analyzeBtn) {
        analyzeBtn.disabled = show;
        analyzeBtn.textContent = show ? "⏳ Analiz Ediliyor..." : "📊 Analiz Başlat";
        analyzeBtn.textContent = show ? "⏳ Analiz Ediliyor..." : "📊 Otomatik Rapor Oluştur";
    }
}
function showResult(text) {
    const resultArea = document.getElementById("resultArea");
    const resultText = document.getElementById("resultText");
    if (resultArea && resultText) {
        resultText.textContent = text;
        resultArea.classList.remove("hidden");
    }
}
function hideResult() { document.getElementById("resultArea")?.classList.add("hidden"); }
function showError(message) {
    const errorArea = document.getElementById("errorArea");
    const errorText = document.getElementById("errorText");
    if (errorArea && errorText) {
        errorText.textContent = message;
        errorArea.classList.remove("hidden");
    }
}
function hideError() { document.getElementById("errorArea")?.classList.add("hidden"); }
