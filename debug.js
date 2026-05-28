require('dotenv').config();
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');
const { GoogleGenerativeAI } = require('@google/generative-ai');

function getAuth() {
    if (!fs.existsSync('./service_account.json')) return null;
    const creds = require('./service_account.json');
    return new JWT({ email: creds.client_email, key: creds.private_key, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
}

function getKategori(keterangan) {
    const text = keterangan.toLowerCase();
    if (/(makan|minum|kopi|resto|gofood|grabfood|jajan|indomaret|alfamart|mcd|kfc)/.test(text)) return 'Makan & Minum';
    if (/(bensin|tol|parkir|ojek|gojek|grab|kereta|pesawat|tiket|travel|transport)/.test(text)) return 'Transportasi';
    if (/(listrik|air|wifi|internet|pulsa|kos|kontrakan|cicilan|asuransi|pdam|tagihan)/.test(text)) return 'Tagihan & Cicilan';
    if (/(belanja|baju|sepatu|skincare|shopee|tokped|tokopedia|lazada|supermarket|pasar)/.test(text)) return 'Belanja';
    if (/(nonton|bioskop|game|main|netflix|spotify|liburan|olahraga|gym|futsal|badminton|renang)/.test(text)) return 'Hiburan & Olahraga';
    return 'Lain-lain';
}

async function updateDashboard(doc, serviceAccountAuth, currentMonthTitle) {
    // Copy-pasting logic exactly from index.js
    let allTimeMasuk = 0; let allTimeKeluar = 0; let monthMasuk = 0; let monthKeluar = 0; let catTotals = {};

    for (const sheet of doc.sheetsByIndex) {
        if (sheet.title === '📊 Dashboard Utama' || sheet.title === '⏰ Pengingat' || sheet.title === '🎯 Anggaran') continue;
        
        const rows = await sheet.getRows();
        for (const row of rows) {
            const tipe = row.get('Tipe');
            const jumlahStr = row.get('Jumlah') || '0';
            const jumlah = parseFloat(jumlahStr.toString().replace(/[^\d.-]/g, ''));
            const isIncome = tipe === '📈 Pemasukan' || tipe === 'pemasukan';
            
            if (isIncome) allTimeMasuk += jumlah;
            else allTimeKeluar += jumlah;

            if (sheet.title === currentMonthTitle) {
                if (isIncome) monthMasuk += jumlah;
                else {
                    monthKeluar += jumlah;
                    const cat = row.get('Kategori') || 'Lain-lain';
                    catTotals[cat] = (catTotals[cat] || 0) + jumlah;
                }
            }
        }
    }

    let dashboardSheet = doc.sheetsByTitle['📊 Dashboard Utama'];
    let isNew = false;
    
    if (!dashboardSheet) {
        dashboardSheet = await doc.addSheet({ title: '📊 Dashboard Utama', index: 0, gridProperties: { columnCount: 15, rowCount: 40 } });
        isNew = true;
    } else if (dashboardSheet.index !== 0) await dashboardSheet.updateProperties({ index: 0 });

    await dashboardSheet.loadCells('A1:C30');

    dashboardSheet.getCell(0, 0).value = `Ringkasan Bulan Ini (${currentMonthTitle})`;
    dashboardSheet.getCell(0, 0).textFormat = { bold: true, fontSize: 12 };
    dashboardSheet.getCell(1, 0).value = "Bulan"; dashboardSheet.getCell(1, 1).value = currentMonthTitle;
    dashboardSheet.getCell(2, 0).value = "Pemasukan"; dashboardSheet.getCell(2, 1).value = monthMasuk;
    dashboardSheet.getCell(3, 0).value = "Pengeluaran"; dashboardSheet.getCell(3, 1).value = monthKeluar;
    dashboardSheet.getCell(4, 0).value = "Saldo"; dashboardSheet.getCell(4, 1).value = monthMasuk - monthKeluar;

    dashboardSheet.getCell(6, 0).value = "Ringkasan Semua Bulan";
    dashboardSheet.getCell(6, 0).textFormat = { bold: true, fontSize: 12 };
    dashboardSheet.getCell(7, 0).value = "Total Pemasukan"; dashboardSheet.getCell(7, 1).value = allTimeMasuk;
    dashboardSheet.getCell(8, 0).value = "Total Pengeluaran"; dashboardSheet.getCell(8, 1).value = allTimeKeluar;
    dashboardSheet.getCell(9, 0).value = "Sisa Saldo"; dashboardSheet.getCell(9, 1).value = allTimeMasuk - allTimeKeluar;

    dashboardSheet.getCell(11, 0).value = "Distribusi Kategori Bulan Ini";
    dashboardSheet.getCell(11, 0).textFormat = { bold: true, fontSize: 12 };
    
    const categories = Object.keys(catTotals);
    let rowIndex = 12;
    for (const cat of categories) {
        dashboardSheet.getCell(rowIndex, 0).value = cat;
        dashboardSheet.getCell(rowIndex, 1).value = catTotals[cat];
        rowIndex++;
    }
    for (let i = rowIndex; i < 25; i++) {
        dashboardSheet.getCell(i, 0).value = null; dashboardSheet.getCell(i, 1).value = null;
    }

    await dashboardSheet.saveUpdatedCells();
}

async function getSheet(doc, serviceAccountAuth) {
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const date = new Date();
    const currentMonthTitle = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
    let sheet = doc.sheetsByTitle[currentMonthTitle];
    
    if (!sheet) {
        sheet = await doc.addSheet({ title: currentMonthTitle, headerValues: ['Tanggal', 'Tipe', 'Jumlah', 'Keterangan', 'Kategori'], gridProperties: { frozenRowCount: 1 } });
    } else {
        try {
            await sheet.loadHeaderRow();
            if (sheet.headerValues.length === 4) await sheet.setHeaderRow(['Tanggal', 'Tipe', 'Jumlah', 'Keterangan', 'Kategori']);
            
            await sheet.loadCells('A1:E1');
            for (let i = 0; i < 5; i++) {
                const cell = sheet.getCell(0, i);
                cell.textFormat = { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } };
                cell.backgroundColor = { red: 0.1, green: 0.2, blue: 0.4 };
                cell.horizontalAlignment = 'CENTER';
            }
            await sheet.saveUpdatedCells();
        } catch(e) {}
    }
    await updateDashboard(doc, serviceAccountAuth, currentMonthTitle);
    return sheet;
}

async function run() {
    try {
        const auth = getAuth();
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        console.log("Loaded doc:", doc.title);
        const sheet = await getSheet(doc, auth);
        console.log("Got sheet:", sheet.title);
        
        await sheet.addRow({ Tanggal: 'test', Tipe: '📉 Pengeluaran', Jumlah: 50000, Keterangan: 'test', Kategori: 'Tagihan & Cicilan' });
        console.log("Added row");
        
        const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        await updateDashboard(doc, auth, `${monthNames[new Date().getMonth()]} ${new Date().getFullYear()}`);
        console.log("Updated dashboard. SUCCESS!");
    } catch (e) {
        console.error("CRASHED:", e);
    }
}
run();
