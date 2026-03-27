// Office.js hazır olduğunda çalışır
Office.onReady((info) => {
    console.log("Office.js hazır. Host:", info.host);
    
    // Butonu bul
    const analyzeBtn = document.getElementById("analyzeBtn");
    
    if (analyzeBtn) {
        console.log("Buton bulundu, event ekleniyor...");
        analyzeBtn.addEventListener("click", analyzeData);
    } else {
        console.error("Buton bulunamadı! ID kontrol edin.");
        showError("Buton bulunamadı. Lütfen sayfayı yenileyin.");
    }
});

/**
 * Ana analiz fonksiyonu
 */
async function analyzeData() {
    console.log("Butona tıklandı! Analiz başlıyor...");
    
    // UI durumunu güncelle
    showLoading(true);
    hideResult();
    hideError();
    
    try {
        // Excel API'sini çalıştır
        await Excel.run(async (context) => {
            console.log("Excel.run başladı");
            
            // Aktif çalışma sayfasını al
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            
            // Kullanılan tüm hücreleri al
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount, columnCount, address");
            
            // Veriyi al
            await context.sync();
            
            console.log("Sayfa adı:", sheet.name);
            console.log("Veri aralığı:", usedRange.address);
            console.log("Satır sayısı:", usedRange.rowCount);
            console.log("Sütun sayısı:", usedRange.columnCount);
            
            // Veri kontrolü
            if (!usedRange.values || usedRange.values.length === 0) {
                throw new Error("Hiç veri bulunamadı. Lütfen veri içeren bir sayfa seçin.");
            }
            
            // Başlıkları al (ilk satır)
            const headers = usedRange.values[0];
            const dataRows = usedRange.values.slice(1);
            
            // Sayısal verileri bul (Adet sütunu ara)
            let adetIndex = -1;
            headers.forEach((header, idx) => {
                const h = String(header || "").toLowerCase();
                if (h.includes("adet") || h.includes("miktar") || h.includes("quantity")) {
                    adetIndex = idx;
                }
            });
            
            // Toplam adet hesapla
            let totalAdet = 0;
            let adetCount = 0;
            
            if (adetIndex !== -1) {
                for (let i = 0; i < dataRows.length; i++) {
                    const val = parseFloat(dataRows[i][adetIndex]);
                    if (!isNaN(val)) {
                        totalAdet += val;
                        adetCount++;
                    }
                }
            }
            
            // Sonucu göster
            const resultText = `
                ✅ ANALİZ TAMAMLANDI!
                
                📄 Sayfa: ${sheet.name}
                📍 Aralık: ${usedRange.address}
                📊 Satır: ${usedRange.rowCount}
                📈 Sütun: ${usedRange.columnCount}
                
                🔍 Başlıklar: ${headers.join(", ")}
                
                ${adetIndex !== -1 ? `
                📦 Adet Sütunu: ${headers[adetIndex]}
                📊 Toplam Adet: ${totalAdet}
                📋 Adet Sayısı: ${adetCount}
                📈 Ortalama Adet: ${(totalAdet / adetCount).toFixed(2)}
                ` : "⚠️ 'Adet' sütunu bulunamadı. Lütfen bir adet/miktar sütunu ekleyin."}
                
                💡 İpucu: AI_ANALYSIS sayfası oluşturulacak.
            `;
            
            showResult(resultText);
            
            // İleride AI_ANALYSIS sayfası oluşturulacak
            console.log("Analiz başarıyla tamamlandı!");
        });
        
    } catch (error) {
        console.error("Hata oluştu:", error);
        showError("Hata: " + error.message);
    } finally {
        showLoading(false);
    }
}

/**
 * UI Yardımcı Fonksiyonları
 */
function showLoading(show) {
    const loading = document.getElementById("loading");
    const analyzeBtn = document.getElementById("analyzeBtn");
    
    if (loading) {
        if (show) {
            loading.classList.remove("hidden");
        } else {
            loading.classList.add("hidden");
        }
    }
    
    if (analyzeBtn) {
        analyzeBtn.disabled = show;
        if (show) {
            analyzeBtn.textContent = "⏳ Analiz Ediliyor...";
        } else {
            analyzeBtn.textContent = "📊 Veriyi Analiz Et";
        }
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

function hideResult() {
    const resultArea = document.getElementById("resultArea");
    if (resultArea) {
        resultArea.classList.add("hidden");
    }
}

function showError(message) {
    const errorArea = document.getElementById("errorArea");
    const errorText = document.getElementById("errorText");
    
    if (errorArea && errorText) {
        errorText.textContent = message;
        errorArea.classList.remove("hidden");
    }
}

function hideError() {
    const errorArea = document.getElementById("errorArea");
    if (errorArea) {
        errorArea.classList.add("hidden");
    }
}
