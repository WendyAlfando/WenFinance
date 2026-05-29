require('dotenv').config();
const dns = require('dns');
if (dns.setDefaultResultOrder) dns.setDefaultResultOrder('ipv4first');
const { Telegraf, Markup } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');
const cron = require('node-cron');
const PDFDocument = require('pdfkit');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const express = require('express');
const cors = require('cors');
const axios = require('axios');

// 1. Inisialisasi Bot Telegram & AI
const bot = new Telegraf(process.env.TELEGRAM_TOKEN);
const genAI = process.env.GEMINI_API_KEY ? new GoogleGenerativeAI(process.env.GEMINI_API_KEY) : null;
const app = express();
app.use(cors());
app.use(express.static('public'));

// 2. Auth Helper
function getAuth() {
    if (!fs.existsSync('./service_account.json')) {
        console.error("service_account.json tidak ditemukan!");
        return null;
    }
    const creds = require('./service_account.json');
    return new JWT({ email: creds.client_email, key: creds.private_key, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
}

// 3. Helper Kategori
function getKategori(keterangan) {
    const text = keterangan.toLowerCase();
    if (/(makan|minum|kopi|resto|gofood|grabfood|jajan|indomaret|alfamart|mcd|kfc)/.test(text)) return 'Makan & Minum';
    if (/(bensin|tol|parkir|ojek|gojek|grab|kereta|pesawat|tiket|travel|transport)/.test(text)) return 'Transportasi';
    if (/(listrik|air|wifi|internet|pulsa|kos|kontrakan|cicilan|asuransi|pdam|tagihan)/.test(text)) return 'Tagihan & Cicilan';
    if (/(belanja|baju|sepatu|skincare|shopee|tokped|tokopedia|lazada|supermarket|pasar)/.test(text)) return 'Belanja';
    if (/(nonton|bioskop|game|main|netflix|spotify|liburan|olahraga|gym|futsal|badminton|renang)/.test(text)) return 'Hiburan & Olahraga';
    return 'Lain-lain';
}

// 4. Budget Helpers
async function checkBudget(doc, currentMonthTitle, kategori, tambahan) {
    let budgetSheet = doc.sheetsByTitle['🎯 Anggaran'];
    if (!budgetSheet) return null;
    
    let batas = 0;
    const budgetRows = await budgetSheet.getRows();
    for (const r of budgetRows) {
        if (r.get('Kategori').toLowerCase() === kategori.toLowerCase()) {
            batas = parseFloat(r.get('Batas Anggaran'));
            break;
        }
    }
    if (batas <= 0) return null;

    let currentTotal = tambahan;
    const monthSheet = doc.sheetsByTitle[currentMonthTitle];
    if (monthSheet) {
        const rows = await monthSheet.getRows();
        for (const row of rows) {
            const tipe = row.get('Tipe');
            if (tipe === '📉 Pengeluaran' || tipe === 'pengeluaran') {
                if ((row.get('Kategori') || 'Lain-lain').toLowerCase() === kategori.toLowerCase()) {
                    currentTotal += parseFloat(row.get('Jumlah').toString().replace(/[^\d.-]/g, ''));
                }
            }
        }
    }
    
    if (currentTotal > batas) return { batas, currentTotal };
    return null;
}

// 5. Fungsi Update Dashboard Google Sheets
async function updateDashboard(doc, serviceAccountAuth, currentMonthTitle) {
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
                if (isIncome) {
                    monthMasuk += jumlah;
                    const cat = row.get('Kategori') || 'Pemasukan';
                    catTotals[cat] = (catTotals[cat] || 0) + jumlah;
                } else {
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
    } else if (dashboardSheet.index !== 0) {
        await dashboardSheet.updateProperties({ index: 0 });
    }

    if (dashboardSheet.rowCount < 30 || dashboardSheet.columnCount < 3) {
        await dashboardSheet.resize({ rowCount: Math.max(dashboardSheet.rowCount, 40), columnCount: Math.max(dashboardSheet.columnCount, 15) }).catch(()=>{});
    }

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

    dashboardSheet.getCell(11, 0).value = "Distribusi Keuangan Bulan Ini";
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

    if (isNew) {
        const sheetId = dashboardSheet.sheetId;
        try {
            await serviceAccountAuth.request({
                method: 'POST',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${doc.spreadsheetId}:batchUpdate`,
                data: {
                    requests: [
                        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 180 }, fields: 'pixelSize' } },
                        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },
                        { repeatCell: { range: { sheetId, startColumnIndex: 1, endColumnIndex: 2 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"Rp"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } }
                    ]
                }
            });
        } catch (err) {}
    }
}

// 6. Fungsi mengakses sheet bulan ini
async function getSheet(doc, serviceAccountAuth) {
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const date = new Date();
    const currentMonthTitle = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
    let sheet = doc.sheetsByTitle[currentMonthTitle];
    
    if (!sheet) {
        sheet = await doc.addSheet({ title: currentMonthTitle, headerValues: ['Tanggal', 'Tipe', 'Jumlah', 'Keterangan', 'Kategori'], gridProperties: { frozenRowCount: 1 } });
        await sheet.loadCells('A1:E1');
        for (let i = 0; i < 5; i++) {
            const cell = sheet.getCell(0, i);
            cell.textFormat = { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } };
            cell.backgroundColor = { red: 0.1, green: 0.2, blue: 0.4 };
            cell.horizontalAlignment = 'CENTER';
        }
        await sheet.saveUpdatedCells();
        try {
            await serviceAccountAuth.request({
                method: 'POST', url: `https://sheets.googleapis.com/v4/spreadsheets/${doc.spreadsheetId}:batchUpdate`,
                data: { requests: [
                    { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },
                    { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
                    { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
                    { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 250 }, fields: 'pixelSize' } },
                    { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },
                    { repeatCell: { range: { sheetId: sheet.sheetId, startRowIndex: 1, startColumnIndex: 2, endColumnIndex: 3 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"Rp"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } }
                ] }
            });
        } catch (err) {}
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

// 7. Fungsi Pengingat Bulanan
async function addReminder(doc, ket, tgl, chatId) {
    let sheet = doc.sheetsByTitle['⏰ Pengingat'];
    if (!sheet) sheet = await doc.addSheet({ title: '⏰ Pengingat', headerValues: ['Tanggal', 'Keterangan', 'ChatID'], gridProperties: { frozenRowCount: 1 } });
    await sheet.addRow({ Tanggal: tgl, Keterangan: ket, ChatID: chatId });
}

cron.schedule('0 8 * * *', async () => {
    const auth = getAuth();
    if (!auth) return;
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['⏰ Pengingat'];
        if (!sheet) return;
        const rows = await sheet.getRows();
        const today = new Date().getDate();
        for (const row of rows) {
            if (parseInt(row.get('Tanggal')) === today) {
                const chatId = row.get('ChatID'); const ket = row.get('Keterangan');
                if (chatId && ket) bot.telegram.sendMessage(chatId, `⏰ *PENGINGAT TAGIHAN*\n\nHalo! Hari ini waktunya: **${ket}**.\nJika sudah dibayar, catat pengeluarannya ya!`, { parse_mode: 'Markdown' }).catch(()=>{});
            }
        }

        // Cek Langganan
        const subSheet = doc.sheetsByTitle['🔄 Langganan'];
        if (subSheet) {
            const subRows = await subSheet.getRows();
            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            const tmrwDate = tomorrow.getDate();
            for (const row of subRows) {
                if (parseInt(row.get('Tgl Jatuh Tempo')) === tmrwDate) {
                    const chatId = row.get('ChatID'); 
                    const ket = row.get('Nama Layanan');
                    const harga = row.get('Harga');
                    if (chatId && ket) {
                        bot.telegram.sendMessage(chatId, `⚠️ *PENGINGAT LANGGANAN H-1*\n\nBesok waktunya bayar **${ket}** sebesar Rp ${harga}.\nSilakan klik tombol di bawah jika sudah dibayar!`, { 
                            parse_mode: 'Markdown',
                            reply_markup: { inline_keyboard: [[ { text: '✅ Sudah Dibayar', callback_data: `pay_sub_${ket}_${harga}` } ]] }
                        }).catch(()=>{});
                    }
                }
            }
        }
    } catch(e) { console.error("Cron Error", e); }
}, { timezone: "Asia/Jakarta" });

// 8. Menu Telegram
const mainMenu = Markup.inlineKeyboard([
    [Markup.button.callback('📊 Ringkasan Bulanan', 'btn_ringkasan'), Markup.button.callback('📈 Grafik Harian', 'btn_grafik_harian')],
    [Markup.button.callback('📄 Export Laporan PDF', 'btn_laporan'), Markup.button.callback('🔙 Batal Terakhir', 'btn_undo')],
    [Markup.button.url('🌐 Buka Dashboard Visual', 'http://146.190.85.119:3000')]
]);

bot.start(async (ctx) => {
    await ctx.reply('Memperbarui sistem menu...', Markup.removeKeyboard());
    ctx.reply(
        'Halo! 👋 Saya adalah Bot Pencatat Keuangan Cerdas.\n\n' +
        '🤖 *AI Mode:* Ketik santai "tadi siang makan ayam 50rb"!\n' +
        '🟢 *Manual Pemasukan:* `+15000000 gaji`\n' +
        '🔴 *Manual Pengeluaran:* `50000 makan`\n' +
        '⏰ *Pengingat:* Tambahkan `tgl <angka>` (Contoh: `wifi tgl 20`)\n' +
        '🎯 *Budget:* Ketik `!budget Makan 1000000`\n\n' +
        'Gunakan tombol di bawah 👇',
        { parse_mode: 'Markdown', ...mainMenu }
    );
});

bot.command('menu', async (ctx) => {
    await ctx.reply('Memperbarui menu...', Markup.removeKeyboard());
    ctx.reply('🎛 *Menu Utama*', { parse_mode: 'Markdown', ...mainMenu });
});

bot.hears(['📊 Ringkasan Keuangan', '💡 Bantuan', '🔙 Batal Terakhir (Undo)', '📊 Ringkasan Bulan Ini'], async (ctx) => {
    await ctx.reply('Sistem menu telah diperbarui. Menghapus menu lama...', Markup.removeKeyboard());
    ctx.reply('Gunakan tombol menu interaktif yang baru di bawah ini 👇', { parse_mode: 'Markdown', ...mainMenu });
});

bot.hears(/^!budget\s+(.+?)\s+(\d+)$/i, async (ctx) => {
    const kategori = ctx.match[1].trim();
    const nominal = parseInt(ctx.match[2], 10);
    const loadingMsg = await ctx.reply('⏳ Menyimpan anggaran...');

    const auth = getAuth();
    if (!auth) return ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Auth error');
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    await doc.loadInfo();

    let sheet = doc.sheetsByTitle['🎯 Anggaran'];
    if (!sheet) sheet = await doc.addSheet({ title: '🎯 Anggaran', headerValues: ['Kategori', 'Batas Anggaran'] });
    
    const rows = await sheet.getRows();
    let found = false;
    for (const r of rows) {
        if (r.get('Kategori').toLowerCase() === kategori.toLowerCase()) {
            r.assign({ 'Batas Anggaran': nominal });
            await r.save();
            found = true; break;
        }
    }
    if (!found) await sheet.addRow({ Kategori: kategori, 'Batas Anggaran': nominal });

    ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, `✅ Anggaran untuk **${kategori}** berhasil diatur menjadi **Rp ${nominal.toLocaleString('id-ID')}**!`, { parse_mode: 'Markdown' });
});

bot.action('btn_ringkasan', async (ctx) => {
    ctx.answerCbQuery('Menghitung ringkasan & membuat grafik...');
    const serviceAccountAuth = getAuth();
    if (!serviceAccountAuth) return ctx.reply('❌ service_account.json hilang.');
    
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = await getSheet(doc, serviceAccountAuth);

    const rows = await sheet.getRows();
    let catTotals = {};
    for (const row of rows) {
        if (row.get('Tipe').includes('Pengeluaran')) {
            const cat = row.get('Kategori') || 'Lain-lain';
            const val = parseFloat(row.get('Jumlah').replace(/[^\d.-]/g, ''));
            catTotals[cat] = (catTotals[cat] || 0) + val;
        }
    }

    const labels = Object.keys(catTotals);
    const data = Object.values(catTotals);
    
    if (labels.length > 0) {
        const chartConfig = {
            type: 'outlabeledPie',
            data: { labels: labels, datasets: [{ data: data }] },
            options: { plugins: { legend: false, outlabels: { text: '%l: Rp %v', color: 'white', stretch: 35, font: { resizable: true, minSize: 12 } } } }
        };
        const chartUrl = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(chartConfig))}&w=600&h=400`;
        await ctx.replyWithPhoto(chartUrl, { caption: '📊 Grafik Distribusi Pengeluaran Bulan Ini', ...mainMenu });
    } else {
        ctx.reply('✅ Belum ada pengeluaran bulan ini.', { parse_mode: 'Markdown', ...mainMenu });
    }
});

bot.action('btn_help', (ctx) => { ctx.answerCbQuery(); ctx.reply('📘 Gunakan AI atau manual!', { parse_mode: 'Markdown' }); });
bot.action('btn_undo', async (ctx) => { /* logic sama seperti sebelumnya, skip implementasi agar rapi, implement basic undo*/
    ctx.answerCbQuery('Menghapus data terakhir...');
    const auth = getAuth();
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    await doc.loadInfo(); const sheet = await getSheet(doc, auth);
    const rows = await sheet.getRows();
    if(rows.length>0) {
        const lastRow = rows[rows.length - 1];
        await lastRow.delete();
        ctx.reply('✅ Transaksi terakhir dibatalkan.');
    } else ctx.reply('⚠️ Kosong.');
});
bot.action('btn_grafik_harian', async (ctx) => {
    ctx.answerCbQuery('Membuat grafik harian...');
    const auth = getAuth();
    if (!auth) return;
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    await doc.loadInfo();
    const sheet = await getSheet(doc, auth);
    const rows = await sheet.getRows();
    
    let dailyTotals = {};
    for (const row of rows) {
        if (row.get('Tipe').includes('Pengeluaran')) {
            // Tanggal format: "DD/MM/YYYY, HH:mm" -> get "DD"
            const dateStr = row.get('Tanggal') || '';
            const dayMatch = dateStr.match(/^(\d{2})\//);
            if (dayMatch) {
                const day = dayMatch[1];
                const val = parseFloat(row.get('Jumlah').replace(/[^\d.-]/g, ''));
                dailyTotals[day] = (dailyTotals[day] || 0) + val;
            }
        }
    }
    
    const labels = Object.keys(dailyTotals).sort((a,b) => parseInt(a) - parseInt(b));
    const data = labels.map(l => dailyTotals[l]);
    
    if (labels.length > 0) {
        const chartConfig = {
            type: 'bar',
            data: { labels: labels.map(l => 'Tgl ' + l), datasets: [{ label: 'Pengeluaran', data: data, backgroundColor: 'rgba(255, 99, 132, 0.8)' }] }
        };
        const chartUrl = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(chartConfig))}&w=600&h=400`;
        await ctx.replyWithPhoto(chartUrl, { caption: '📈 Grafik Tren Pengeluaran Harian', ...mainMenu });
    } else {
        ctx.reply('✅ Belum ada pengeluaran bulan ini untuk grafik.', { ...mainMenu });
    }
});

bot.action('btn_laporan', async (ctx) => {
    const loadingMsg = await ctx.reply('⏳ Sedang men-generate file PDF...');
    ctx.answerCbQuery();
    try {
        const auth = getAuth();
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        const sheet = await getSheet(doc, auth);
        const rows = await sheet.getRows();
        
        let totalPemasukan = 0;
        let totalPengeluaran = 0;
        const transactions = [];

        for (const row of rows) {
            const tipe = row.get('Tipe');
            const jumlah = parseFloat(row.get('Jumlah').replace(/[^\d.-]/g, ''));
            const isIncome = tipe.includes('Pemasukan');
            if (isIncome) totalPemasukan += jumlah;
            else totalPengeluaran += jumlah;
            
            transactions.push({
                tgl: row.get('Tanggal').split(',')[0],
                tipe: isIncome ? '+' : '-',
                jumlah: jumlah,
                ket: row.get('Keterangan')
            });
        }
        
        // Buat PDF
        const pdf = new PDFDocument({ margin: 50 });
        const buffers = [];
        pdf.on('data', buffers.push.bind(buffers));
        
        // Isi konten PDF
        pdf.fontSize(20).text('Laporan Keuangan Bulanan', { align: 'center' });
        pdf.moveDown();
        pdf.fontSize(14).text(`Bulan: ${sheet.title}`);
        pdf.text(`Total Pemasukan: Rp ${totalPemasukan.toLocaleString('id-ID')}`);
        pdf.text(`Total Pengeluaran: Rp ${totalPengeluaran.toLocaleString('id-ID')}`);
        pdf.text(`Saldo Akhir: Rp ${(totalPemasukan - totalPengeluaran).toLocaleString('id-ID')}`);
        pdf.moveDown(2);
        
        pdf.fontSize(12).text('Rincian Transaksi:', { underline: true });
        pdf.moveDown(0.5);
        for (const tx of transactions) {
            const color = tx.tipe === '+' ? 'green' : 'red';
            pdf.fillColor(color).text(`${tx.tgl} | ${tx.tipe} Rp ${tx.jumlah.toLocaleString('id-ID')} | ${tx.ket}`);
        }
        
        pdf.end();
        pdf.on('end', async () => {
            const pdfData = Buffer.concat(buffers);
            await ctx.replyWithDocument({ source: pdfData, filename: `Laporan_${sheet.title.replace(' ', '_')}.pdf` });
            await ctx.telegram.deleteMessage(ctx.chat.id, loadingMsg.message_id).catch(()=>{});
        });
        
    } catch (err) {
        console.error(err);
        await ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Gagal men-generate PDF.');
    }
});

bot.on('text', async (ctx) => {
    try {
        const text = ctx.message.text.trim();
        if (text.startsWith('/')) return;

        let transactions = [];
        const match = text.match(/^(\+?)(\d+)\s+(.+)$/);

        if (match) {
            const isPemasukan = match[1] === '+';
            const jumlah = parseInt(match[2], 10);
            let keterangan = match[3].trim();
            const tipe = isPemasukan ? '📈 Pemasukan' : '📉 Pengeluaran';
            if (jumlah <= 0) return await ctx.reply('⚠️ Jumlah uang harus > 0.');

            let reminderDate = null;
            const remMatch = keterangan.match(/\s+(?:tgl|tanggal|tiap)\s*(\d{1,2})$/i);
            if (remMatch) {
                const parsed = parseInt(remMatch[1], 10);
                if (parsed >= 1 && parsed <= 31) { reminderDate = parsed; keterangan = keterangan.replace(remMatch[0], '').trim(); }
            }
            transactions.push({ tipe, jumlah, keterangan, kategori: isPemasukan ? 'Pemasukan' : getKategori(keterangan), dompet: 'Cash', reminderDate });
        } else {
            if (!genAI) return await ctx.reply('❌ Format tidak dikenali dan GEMINI API KEY tidak ada.\n\nKetik: `50000 kopi`', { parse_mode: 'Markdown' });
            const loadingAI = await ctx.reply('🤖 Memproses dengan AI...');
            try {
                const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
                const prompt = `Ekstrak satu atau beberapa transaksi keuangan dari teks ini: "${text}". 
Deteksi juga "dompet" atau rekening asal (BCA, Mandiri, Gopay, Dana, OVO, ShopeePay, atau Cash). Jika tidak disebutkan, default ke "Cash".
Output HANYA array JSON murni: [{"tipe": "Pengeluaran"|"Pemasukan", "jumlah": <angka>, "keterangan": "<deskripsi singkat>", "dompet": "<Nama Dompet>"}] tanpa penjelasan lain.`;
                const result = await model.generateContent(prompt);
                
                let rawText = result.response.text();
                // Cari blok JSON (array atau object) di dalam text yang mungkin cerewet
                let start = rawText.indexOf('[');
                let end = rawText.lastIndexOf(']');
                if (start === -1) {
                    start = rawText.indexOf('{');
                    end = rawText.lastIndexOf('}');
                }
                
                if (start === -1 || end === -1) throw new Error("Tidak ada JSON ditemukan: " + rawText);
                
                let jsonText = rawText.substring(start, end + 1);
                const parsed = JSON.parse(jsonText);
                
                // Fix for when Gemini returns an object instead of array
                const parsedArray = Array.isArray(parsed) ? parsed : [parsed];
                
                for (const item of parsedArray) {
                    transactions.push({
                        tipe: item.tipe === 'Pemasukan' ? '📈 Pemasukan' : '📉 Pengeluaran',
                        jumlah: item.jumlah, keterangan: item.keterangan,
                        kategori: item.tipe === 'Pemasukan' ? 'Pemasukan' : getKategori(item.keterangan),
                        dompet: item.dompet || 'Cash',
                        reminderDate: null
                    });
                }
                await ctx.telegram.deleteMessage(ctx.chat.id, loadingAI.message_id).catch(() => {});
            } catch (e) {
                console.error("AI Parse Error:", e);
                return await ctx.telegram.editMessageText(ctx.chat.id, loadingAI.message_id, null, `❌ AI Error: ${e.message}`).catch(() => {});
            }
        }

        if (transactions.length === 0) return;

        const loadingMsg = await ctx.reply('⏳ Menyimpan...');
        try {
            const serviceAccountAuth = getAuth();
            const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
            await doc.loadInfo();
            const sheet = await getSheet(doc, serviceAccountAuth);
            const dateStr = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });

            let finalReply = "✅ *Transaksi Berhasil Dicatat!*\n";
            for (const tx of transactions) {
                await sheet.addRow({ Tanggal: dateStr, Tipe: tx.tipe, Jumlah: tx.jumlah, Keterangan: tx.keterangan, Kategori: tx.kategori, Dompet: tx.dompet || 'Cash' });
                if (tx.reminderDate) await addReminder(doc, tx.keterangan, tx.reminderDate, ctx.chat.id);

                let alertMsg = "";
                if (tx.tipe === '📉 Pengeluaran') {
                    const currentMonthTitle = `${["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][new Date().getMonth()]} ${new Date().getFullYear()}`;
                    const budgetRes = await checkBudget(doc, currentMonthTitle, tx.kategori, tx.jumlah);
                    if (budgetRes) alertMsg = `\n⚠️ *AWAS BUDGET OVERLOAD!*\nKategori *${tx.kategori}* melampaui batas (Rp ${budgetRes.currentTotal.toLocaleString('id-ID')} / Rp ${budgetRes.batas.toLocaleString('id-ID')})!`;
                }

                const icon = tx.tipe === '📉 Pengeluaran' ? '📉' : '📈';
                finalReply += `\n${icon} Rp ${tx.jumlah.toLocaleString('id-ID')} | ${tx.keterangan}`;
                if (tx.reminderDate) finalReply += ` *(Ingat tgl ${tx.reminderDate})*`;
                if (alertMsg) finalReply += alertMsg;
            }

            const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
            await updateDashboard(doc, serviceAccountAuth, `${monthNames[new Date().getMonth()]} ${new Date().getFullYear()}`);
            
            // Hapus pesan loading dan kirim pesan baru agar muncul notifikasi push
            await ctx.telegram.deleteMessage(ctx.chat.id, loadingMsg.message_id).catch(() => {});
            await ctx.reply(finalReply, { parse_mode: 'Markdown', ...mainMenu }).catch(() => {});
        } catch (err) {
            console.error(err);
            await ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, `❌ Kesalahan Sistem: ${err.message}`).catch(() => {});
        }
    } catch (fatalError) {
        console.error("FATAL ERROR:", fatalError);
        await ctx.reply(`❌ FATAL ERROR: ${fatalError.message}`).catch(() => {});
    }
});

bot.action(/pay_sub_(.+)_(.+)/, async (ctx) => {
    const ket = ctx.match[1];
    const jumlah = parseInt(ctx.match[2]);
    const loadingMsg = await ctx.reply('⏳ Mencatat...');
    ctx.answerCbQuery();
    try {
        const auth = getAuth();
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        const sheet = await getSheet(doc, auth);
        const dateStr = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
        await sheet.addRow({ Tanggal: dateStr, Tipe: '📉 Pengeluaran', Jumlah: jumlah, Keterangan: ket, Kategori: 'Tagihan & Cicilan', Dompet: 'Cash' });
        await ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, `✅ Pembayaran **${ket}** (Rp ${jumlah.toLocaleString('id-ID')}) dicatat!`, {parse_mode:'Markdown'});
    } catch(e) { await ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Gagal mencatat.'); }
});

bot.on('photo', async (ctx) => {
    if (!genAI) return ctx.reply('❌ GEMINI API KEY tidak ada. Fitur scan struk tidak bisa digunakan.');
    const loading = await ctx.reply('📸 Membaca struk dengan AI...');
    try {
        const photo = ctx.message.photo[ctx.message.photo.length - 1];
        const fileLink = await ctx.telegram.getFileLink(photo.file_id);
        const response = await axios.get(fileLink.href, { responseType: 'arraybuffer' });
        const base64Data = Buffer.from(response.data, 'binary').toString('base64');
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const prompt = `Ini foto struk. Output HANYA JSON: {"jumlah": <angka_total_tanpa_titik_tanpa_rp>, "keterangan": "Belanja di <Nama Toko>"}`;
        const result = await model.generateContent([{ inlineData: { data: base64Data, mimeType: 'image/jpeg' } }, prompt]);
        let rawText = result.response.text();
        let start = rawText.indexOf('{');
        let end = rawText.lastIndexOf('}');
        const parsed = JSON.parse(rawText.substring(start, end + 1));
        const auth = getAuth();
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        const sheet = await getSheet(doc, auth);
        const dateStr = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
        await sheet.addRow({ Tanggal: dateStr, Tipe: '📉 Pengeluaran', Jumlah: parsed.jumlah, Keterangan: parsed.keterangan, Kategori: getKategori(parsed.keterangan), Dompet: 'Cash' });
        await ctx.telegram.deleteMessage(ctx.chat.id, loading.message_id).catch(()=>{});
        ctx.reply(`✅ **Struk Berhasil Dicatat!**\n📉 Rp ${parseInt(parsed.jumlah).toLocaleString('id-ID')} | ${parsed.keterangan}`, {parse_mode: 'Markdown'});
    } catch(e) { await ctx.telegram.editMessageText(ctx.chat.id, loading.message_id, null, `❌ Gagal membaca struk: ${e.message}`); }
});

app.get('/api/dashboard', async (req, res) => {
    try {
        const auth = getAuth();
        if (!auth) return res.status(500).json({error: 'Auth failed'});
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        const sheet = await getSheet(doc, auth);
        const rows = await sheet.getRows();
        let totalMasuk = 0; let totalKeluar = 0; let catTotals = {}; let dailyTotals = {}; let transactions = [];
        let wallets = { 'BCA':0, 'Mandiri':0, 'Gopay':0, 'Dana':0, 'OVO':0, 'ShopeePay':0, 'Cash':0 };
        for (let i = rows.length - 1; i >= Math.max(0, rows.length - 15); i--) {
            const r = rows[i];
            transactions.push({ tgl: r.get('Tanggal').split(',')[0], ket: r.get('Keterangan'), kat: r.get('Kategori'), tipe: r.get('Tipe'), jumlah: parseFloat(r.get('Jumlah').replace(/[^\d.-]/g, '')) });
        }
        for (const row of rows) {
            const tipe = row.get('Tipe');
            const jumlah = parseFloat(row.get('Jumlah').replace(/[^\d.-]/g, ''));
            const dompet = row.get('Dompet') || 'Cash';
            if (wallets[dompet] !== undefined) {
                if (tipe.includes('Pemasukan')) wallets[dompet] += jumlah;
                else wallets[dompet] -= jumlah;
            } else {
                wallets[dompet] = tipe.includes('Pemasukan') ? jumlah : -jumlah;
            }

            if (tipe.includes('Pemasukan')) totalMasuk += jumlah;
            else {
                totalKeluar += jumlah;
                catTotals[row.get('Kategori') || 'Lain-lain'] = (catTotals[row.get('Kategori') || 'Lain-lain'] || 0) + jumlah;
                const dateStr = row.get('Tanggal') || '';
                const dayMatch = dateStr.match(/^(\d{2})\//);
                if (dayMatch) dailyTotals[dayMatch[1]] = (dailyTotals[dayMatch[1]] || 0) + jumlah;
            }
        }
        const categories = Object.keys(catTotals).map(k => ({ name: k, amount: catTotals[k] })).sort((a,b)=>b.amount-a.amount);
        const dailyLabels = Object.keys(dailyTotals).sort((a,b)=>parseInt(a)-parseInt(b));
        const dailyData = dailyLabels.map(l => dailyTotals[l]);
        res.json({ month: sheet.title, totalMasuk, totalKeluar, categories, dailyLabels, dailyData, transactions, wallets });
    } catch(e) { res.status(500).json({error: e.message}); }
});

app.listen(3000, () => console.log('Web App berjalan di port 3000'));
bot.launch().then(() => console.log("Bot Telegram sedang berjalan..."));
process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
