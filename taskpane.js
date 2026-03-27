// Office Add-in hazır olduğunda çalıştır
Office.onReady((info) => {
    console.log("Office.js hazır:", info.host);
    
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (analyzeBtn) {
        analyzeBtn.addEventListener("click", analyzeData);
    }
});

/**
 * Ana analiz fonksiyonu
 */
async function analyzeData() {
    showLoading(true);
    hideResult();
    hideError();
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount, columnCount, address");
            sheet.load("name");
            
            await context.sync();
            
            // Veri kontrolü
            if (!usedRange.values || usedRange.values.length < 2) {
                showError("Veri bulunamadı! En az bir başlık satırı ve bir veri satırı olmalı.");
                showLoading(false);
                return;
            }
            
            const headers = usedRange.values[0];
            const dataRows = usedRange.values.slice(1);
            
            console.log("Başlıklar:", headers);
            console.log("Veri satırları:", dataRows.length);
            
            // Analiz sonuçlarını hesapla
            const analysisResult = analyzeDataRows(dataRows, headers);
            
            // AI_ANALYSIS sayfasını oluştur ve yaz
            await writeAnalysisSheet(context, analysisResult, sheet.name, usedRange.address);
            
            // UI'ı güncelle
            displayResultsUI(usedRange, analysisResult);
            
            await context.sync();
        });
        
    } catch (error) {
        console.error("Hata:", error);
        showError("Analiz sırasında hata: " + error.message);
    } finally {
        showLoading(false);
    }
}

/**
 * Veri satırlarını analiz et
 */
function analyzeDataRows(dataRows, headers) {
    const result = {
        rowCount: dataRows.length,
        colCount: headers.length,
        numericValues: [],
        adetIndex: -1,
        musteriIndex: -1,
        warnings: []
    };
    
    // Kolon indekslerini bul
    headers.forEach((h, idx) => {
        const header = String(h || "").toLowerCase();
        if (header.includes("adet") || header.includes("miktar") || header.includes("quantity")) {
            result.adetIndex = idx;
        }
        if (header.includes("müşteri") || header.includes("musteri") || header.includes("customer") || header.includes("bayi")) {
            result.musteriIndex = idx;
        }
    });
    
    // Adet kolonu bulunamadıysa uyarı ekle
    if (result.adetIndex === -1) {
        result.warnings.push("⚠️ 'Adet' kolonu bulunamadı. Lütfen bir adet/miktar sütunu ekleyin.");
    }
    
    // Sayısal değerleri topla
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        if (result.adetIndex !== -1 && row[result.adetIndex]) {
            const val = parseFloat(row[result.adetIndex]);
            if (!isNaN(val)) {
                result.numericValues.push(val);
            }
        }
    }
    
    // İstatistik hesapla
    if (result.numericValues.length > 0) {
        const sorted = [...result.numericValues].sort((a, b) => a - b);
        const sum = result.numericValues.reduce((a, b) => a + b, 0);
        result.stats = {
            count: result.numericValues.length,
            sum: sum,
            average: sum / result.numericValues.length,
            min: sorted[0],
            max: sorted[sorted.length - 1]
        };
    } else {
        result.stats = { count: 0, sum: 0, average: 0, min: 0, max: 0 };
        if (result.warnings.length === 0) {
            result.warnings.push("⚠️ Sayısal veri bulunamadı. 'Adet' sütununda sayısal değerler olmalı.");
        }
    }
    
    return result;
}

/**
 * AI_ANALYSIS sayfasına rapor yaz
 */
async function writeAnalysisSheet(context, analysis, sheetName, rangeAddress) {
    try {
        // AI_ANALYSIS sayfasını bul veya oluştur
        let analysisSheet;
        const existingSheets = context.workbook.worksheets;
        existingSheets.load("items");
        await context.sync();
        
        const targetName = "AI_ANALYSIS";
        const existingSheet = existingSheets.items.find(s => s.name === targetName);
        
        if (existingSheet) {
            analysisSheet = existingSheet;
            const usedRange = analysisSheet.getUsedRange();
            if (usedRange) {
                usedRange.clear();
            }
        } else {
            analysisSheet = context.workbook.worksheets.add(targetName);
        }
        
        await context.sync();
        
        // Rapor verilerini hazırla (düzenli array olarak)
        const reportLines = [
            ["=== AI ANALİZ RAPORU ==="],
            ["Oluşturulma:", new Date().toLocaleString("tr-TR")],
            [""],
            ["KAYNAK BİLGİLERİ"],
            ["Sayfa:", sheetName],
            ["Veri Aralığı:", rangeAddress],
            ["Satır Sayısı:", analysis.rowCount],
            ["Sütun Sayısı:", analysis.colCount],
            [""],
            ["İSTATİSTİKSEL ANALİZ"],
            ["Sayısal Veri Sayısı:", analysis.stats.count],
            ["Toplam Adet:", analysis.stats.sum],
            ["Ortalama:", analysis.stats.average.toFixed(2)],
            ["Minimum:", analysis.stats.min],
            ["Maksimum:", analysis.stats.max],
            [""],
            ["UYARILAR"]
        ];
        
        // Uyarıları ekle
        if (analysis.warnings.length === 0) {
            reportLines.push(["✅ Herhangi bir uyarı bulunmamaktadır."]);
        } else {
            analysis.warnings.forEach(w => {
                reportLines.push(["⚠️", w]);
            });
        }
        
        reportLines.push([""]);
        reportLines.push(["=== RAPOR SONU ==="]);
        
        // Excel'e yaz - satır ve sütun sayısını doğru hesapla
        const rowCount = reportLines.length;
        const colCount = 2; // Her zaman 2 sütun kullan
        
        const targetRange = analysisSheet.getRangeByIndexes(0, 0, rowCount, colCount);
        
        // reportLines array'ini 2 sütunlu formata dönüştür
        const values = [];
        for (let i = 0; i < reportLines.length; i++) {
            const line = reportLines[i];
            if (line.length === 1) {
                values.push([line[0], ""]);
            } else if (line.length === 2) {
                values.push([line[0], line[1]]);
            } else {
                values.push([line[0], ""]);
            }
        }
        
        targetRange.values = values;
        
        // Sütun genişliklerini ayarla
        targetRange.getColumn(0).format.columnWidth = 30;
        targetRange.getColumn(1).format.columnWidth = 50;
        
        await context.sync();
        
    } catch (error) {
        console.error("Rapor yazma hatası:", error);
        throw error;
    }
}

/**
 * UI'da sonuçları göster
 */
function displayResultsUI(sourceRange, analysis) {
    showResult(true);
    
    const summaryContent = document.getElementById("summaryContent");
    if (summaryContent) {
        summaryContent.innerHTML = `
            <div>📊 ${sourceRange.rowCount} satır × ${sourceRange.columnCount} sütun</div>
            <div>📈 Toplam Adet: ${analysis.stats.sum}</div>
            <div>📊 Ortalama: ${analysis.stats.average.toFixed(2)}</div>
            <div>📉 Min: ${analysis.stats.min} | Max: ${analysis.stats.max}</div>
        `;
    }
    
    const warningsContent = document.getElementById("warningsContent");
    if (warningsContent) {
        if (analysis.warnings.length === 0) {
            warningsContent.innerHTML = '<div class="text-green-600">✅ Sorun tespit edilmedi.</div>';
        } else {
            warningsContent.innerHTML = analysis.warnings.map(w => 
                `<div class="bg-yellow-50 p-2 rounded border-l-4 border-yellow-500">${w}</div>`
            ).join("");
        }
    }
    
    const analysisContent = document.getElementById("analysisContent");
    if (analysisContent) {
        analysisContent.innerHTML = `
            <div class="bg-blue-50 p-3 rounded">
                ✅ Analiz tamamlandı! Detaylı rapor "AI_ANALYSIS" sayfasında.
            </div>
        `;
    }
}

// UI yardımcı fonksiyonları
function showLoading(show) {
    const loadingArea = document.getElementById("loadingArea");
    if (loadingArea) loadingArea.classList.toggle("hidden", !show);
    
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (analyzeBtn) {
        analyzeBtn.disabled = show;
        analyzeBtn.innerHTML = show ? '<span>⏳</span><span>Analiz Ediliyor...</span>' : '<span>🔍</span><span>Veriyi Analiz Et</span>';
    }
}

function showResult(show) {
    const resultArea = document.getElementById("resultArea");
    if (resultArea) resultArea.classList.toggle("hidden", !show);
}

function hideResult() {
    const resultArea = document.getElementById("resultArea");
    if (resultArea) resultArea.classList.add("hidden");
}

function showError(message) {
    const errorArea = document.getElementById("errorArea");
    const errorMessage = document.getElementById("errorMessage");
    if (errorArea && errorMessage) {
        errorMessage.textContent = message;
        errorArea.classList.remove("hidden");
    }
}

function hideError() {
    const errorArea = document.getElementById("errorArea");
    if (errorArea) errorArea.classList.add("hidden");
}
