let cachedData = null;
let cachedAnalysis = null;

Office.onReady(() => {
    console.log("Office hazır");

    document.getElementById("readBtn").onclick = readData;
    document.getElementById("analyzeBtn").onclick = analyze;
    document.getElementById("buildBtn").onclick = buildDashboard;
    document.getElementById("writeBtn").onclick = writeResult;
});

// 1️⃣ VERİ OKU
async function readData() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.load("values");
        await context.sync();

        cachedData = range.values;
        alert("Veri okundu ✔️");
    });
}

// 2️⃣ ANALİZ
function analyze() {
    if (!cachedData) {
        alert("Önce veri oku");
        return;
    }

    const headers = cachedData[0];
    const rows = cachedData.slice(1);

    let total = 0;

    rows.forEach(r => {
        const val = parseFloat(r[2]);
        if (!isNaN(val)) total += val;
    });

    cachedAnalysis = {
        total: total,
        rowCount: rows.length
    };

    alert("Analiz tamam ✔️");
}

// 3️⃣ DASHBOARD KUR
async function buildDashboard() {
    if (!cachedAnalysis) {
        alert("Önce analiz yap");
        return;
    }

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.add("EXECUTIVE_DASHBOARD");

        const data = [
            ["KPI", "Değer"],
            ["Toplam Satış", cachedAnalysis.total],
            ["Satır Sayısı", cachedAnalysis.rowCount]
        ];

        const range = sheet.getRange("A1:B3");
        range.values = data;

        await context.sync();
    });

    alert("Dashboard oluşturuldu 🚀");
}

// 4️⃣ SONUÇ YAZ
async function writeResult() {
    if (!cachedAnalysis) {
        alert("Önce analiz yap");
        return;
    }

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.add("AI_ANALYSIS");

        const data = [
            ["Toplam", cachedAnalysis.total],
            ["Satır", cachedAnalysis.rowCount]
        ];

        sheet.getRange("A1:B2").values = data;

        await context.sync();
    });

    alert("Sonuç yazıldı 📄");
}
