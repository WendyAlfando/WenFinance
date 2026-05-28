require('dotenv').config();
const { Telegraf, Markup } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');
const cron = require('node-cron');
const PDFDocument = require('pdfkit');

// 1. Inisialisasi Bot Telegram
const bot = new Telegraf(process.env.TELEGRAM_TOKEN);

// 2. Auth Helper
function getAuth() {
    if (!fs.existsSync('./service_account.json')) {
        console.error("service_account.json tidak ditemukan!");
        return null;
    }
    const creds = require('./service_account.json');
    return new JWT({
        email: creds.client_email,
        key: creds.private_key,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
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

// 4. Fungsi Update Dashboard Google Sheets
async function updateDashboard(doc, serviceAccountAuth, currentMonthTitle) {
    let allTimeMasuk = 0;
    let allTimeKeluar = 0;
    let monthMasuk = 0;
    let monthKeluar = 0;
    let catTotals = {};

    for (const sheet of doc.sheetsByIndex) {
        if (sheet.title === '📊 Dashboard Utama' || sheet.title === '⏰ Pengingat') continue;
        
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
        dashboardSheet = await doc.addSheet({
            title: '📊 Dashboard Utama',
            index: 0,
            gridProperties: { columnCount: 15, rowCount: 40 }
        });
        isNew = true;
    } else {
        if (dashboardSheet.index !== 0) {
            await dashboardSheet.updateProperties({ index: 0 });
        }
    }

    await dashboardSheet.loadCells('A1:C30');

    // Title 1
    dashboardSheet.getCell(0, 0).value = `Ringkasan Bulan Ini (${currentMonthTitle})`;
    dashboardSheet.getCell(0, 0).textFormat = { bold: true, fontSize: 12 };
    
    dashboardSheet.getCell(1, 0).value = "Bulan";
    dashboardSheet.getCell(1, 1).value = currentMonthTitle;
    dashboardSheet.getCell(2, 0).value = "Pemasukan";
    dashboardSheet.getCell(2, 1).value = monthMasuk;
    dashboardSheet.getCell(3, 0).value = "Pengeluaran";
    dashboardSheet.getCell(3, 1).value = monthKeluar;
    dashboardSheet.getCell(4, 0).value = "Saldo";
    dashboardSheet.getCell(4, 1).value = monthMasuk - monthKeluar;

    // Title 2
    dashboardSheet.getCell(6, 0).value = "Ringkasan Semua Bulan";
    dashboardSheet.getCell(6, 0).textFormat = { bold: true, fontSize: 12 };

    dashboardSheet.getCell(7, 0).value = "Total Pemasukan";
    dashboardSheet.getCell(7, 1).value = allTimeMasuk;
    dashboardSheet.getCell(8, 0).value = "Total Pengeluaran";
    dashboardSheet.getCell(8, 1).value = allTimeKeluar;
    dashboardSheet.getCell(9, 0).value = "Sisa Saldo";
    dashboardSheet.getCell(9, 1).value = allTimeMasuk - allTimeKeluar;

    // Kategori Breakdown
    dashboardSheet.getCell(11, 0).value = "Distribusi Kategori Bulan Ini";
    dashboardSheet.getCell(11, 0).textFormat = { bold: true, fontSize: 12 };
    
    const categories = Object.keys(catTotals);
    let rowIndex = 12;
    for (const cat of categories) {
        dashboardSheet.getCell(rowIndex, 0).value = cat;
        dashboardSheet.getCell(rowIndex, 1).value = catTotals[cat];
        rowIndex++;
    }

    // Bersihkan sel kategori lama jika ada yang berkurang
    for (let i = rowIndex; i < 25; i++) {
        dashboardSheet.getCell(i, 0).value = null;
        dashboardSheet.getCell(i, 1).value = null;
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
                        {
                            repeatCell: {
                                range: { sheetId, startColumnIndex: 1, endColumnIndex: 2 },
                                cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"Rp"#,##0' } } },
                                fields: 'userEnteredFormat.numberFormat'
                            }
                        },
                        {
                            addChart: {
                                chart: {
                                    spec: {
                                        title: "Pengeluaran per Kategori (Bulan Ini)",
                                        pieChart: {
                                            legendPosition: "RIGHT_LEGEND",
                                            domain: { sourceRange: { sources: [{ sheetId, startRowIndex: 12, endRowIndex: 25, startColumnIndex: 0, endColumnIndex: 1 }] } },
                                            series: { sourceRange: { sources: [{ sheetId, startRowIndex: 12, endRowIndex: 25, startColumnIndex: 1, endColumnIndex: 2 }] } }
                                        }
                                    },
                                    position: { overlayPosition: { anchorCell: { sheetId, rowIndex: 12, columnIndex: 3 }, widthPixels: 400, heightPixels: 280 } }
                                }
                            }
                        },
                        {
                            addChart: {
                                chart: {
                                    spec: {
                                        title: "Distribusi Bulan Ini",
                                        pieChart: {
                                            legendPosition: "RIGHT_LEGEND",
                                            domain: { sourceRange: { sources: [{ sheetId, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 1 }] } },
                                            series: { sourceRange: { sources: [{ sheetId, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 1, endColumnIndex: 2 }] } }
                                        }
                                    },
                                    position: { overlayPosition: { anchorCell: { sheetId, rowIndex: 0, columnIndex: 3 }, widthPixels: 400, heightPixels: 250 } }
                                }
                            }
                        },
                        {
                            addChart: {
                                chart: {
                                    spec: {
                                        title: "Distribusi Semua Bulan",
                                        pieChart: {
                                            legendPosition: "RIGHT_LEGEND",
                                            domain: { sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 1 }] } },
                                            series: { sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 9, startColumnIndex: 1, endColumnIndex: 2 }] } }
                                        }
                                    },
                                    position: { overlayPosition: { anchorCell: { sheetId, rowIndex: 0, columnIndex: 9 }, widthPixels: 400, heightPixels: 250 } }
                                }
                            }
                        }
                    ]
                }
            });
        } catch (err) { console.error("Gagal menambah chart:", err); }
    }
}

// 5. Fungsi mengakses sheet bulan ini
async function getSheet(doc, serviceAccountAuth) {
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const date = new Date();
    const currentMonthTitle = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
    
    let sheet = doc.sheetsByTitle[currentMonthTitle];
    
    if (!sheet) {
        sheet = await doc.addSheet({
            title: currentMonthTitle,
            headerValues: ['Tanggal', 'Tipe', 'Jumlah', 'Keterangan', 'Kategori'],
            gridProperties: { frozenRowCount: 1 }
        });

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
                method: 'POST',
                url: `https://sheets.googleapis.com/v4/spreadsheets/${doc.spreadsheetId}:batchUpdate`,
                data: {
                    requests: [
                        { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },
                        { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
                        { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
                        { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 250 }, fields: 'pixelSize' } },
                        { updateDimensionProperties: { range: { sheetId: sheet.sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } },
                        {
                            repeatCell: {
                                range: { sheetId: sheet.sheetId, startRowIndex: 1, startColumnIndex: 2, endColumnIndex: 3 },
                                cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"Rp"#,##0' } } },
                                fields: 'userEnteredFormat.numberFormat'
                            }
                        }
                    ]
                }
            });
        } catch (err) {}
    } else {
        // Upgrade existing sheet headers if needed
        try {
            await sheet.loadHeaderRow();
            if (sheet.headerValues.length === 4) {
                await sheet.setHeaderRow(['Tanggal', 'Tipe', 'Jumlah', 'Keterangan', 'Kategori']);
            }
        } catch(e) {}
    }
    
    await updateDashboard(doc, serviceAccountAuth, currentMonthTitle);
    return sheet;
}

// 6. Fungsi Pengingat Bulanan
async function addReminder(doc, ket, tgl, chatId) {
    let sheet = doc.sheetsByTitle['⏰ Pengingat'];
    if (!sheet) {
        sheet = await doc.addSheet({
            title: '⏰ Pengingat',
            headerValues: ['Tanggal', 'Keterangan', 'ChatID'],
            gridProperties: { frozenRowCount: 1 }
        });
    }
    await sheet.addRow({ Tanggal: tgl, Keterangan: ket, ChatID: chatId });
}

// Cron Job berjalan tiap hari jam 08:00 AM WIB
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
                const chatId = row.get('ChatID');
                const ket = row.get('Keterangan');
                if (chatId && ket) {
                    bot.telegram.sendMessage(chatId, `⏰ *PENGINGAT TAGIHAN*\n\nHalo! Hari ini adalah waktunya untuk membayar/mencatat pengeluaran: **${ket}**.\n\nJika sudah dibayar, catat pengeluarannya di bot ini ya!`, { parse_mode: 'Markdown' }).catch(()=>{});
                }
            }
        }
    } catch(e) { console.error("Cron Error", e); }
}, { timezone: "Asia/Jakarta" });


// 7. Command & Menu Telegram
const mainMenu = Markup.inlineKeyboard([
    [Markup.button.callback('📊 Ringkasan & Dashboard', 'btn_ringkasan'), Markup.button.callback('📄 Export Laporan', 'btn_laporan')],
    [Markup.button.callback('🔙 Batal Terakhir', 'btn_undo'), Markup.button.callback('💡 Bantuan', 'btn_help')]
]);

bot.start(async (ctx) => {
    await ctx.reply('Memperbarui sistem menu...', Markup.removeKeyboard());
    ctx.reply(
        'Halo! 👋 Saya adalah Bot Pencatat Keuangan Anda.\n\n' +
        '🟢 *Pemasukan:* Ketik `+<angka> <keterangan>`\n' +
        '🔴 *Pengeluaran:* Ketik `<angka> <keterangan>`\n' +
        '⏰ *Pengingat:* Tambahkan `tgl <angka>` di akhir ket. (Contoh: `250000 listrik tgl 20`)\n\n' +
        'Gunakan tombol di bawah untuk mengakses menu 👇',
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

bot.action('btn_help', (ctx) => {
    ctx.answerCbQuery();
    ctx.reply(
        '📘 *Cara Menggunakan Bot:*\n\n' +
        '- *Pengeluaran*: `50000 bensin`\n' +
        '- *Pemasukan*: `+2000000 bonus project`\n' +
        '- *Pengingat Otomatis*: Tambahkan kata `tgl` lalu tanggal di akhir pesan. Contoh: `50000 wifi tgl 15` (Bot akan mengingatkanmu tiap tanggal 15 jam 8 pagi).\n' +
        '- *Laporan*: Download PDF/CSV via tombol Export Laporan.',
        { parse_mode: 'Markdown' }
    );
});

bot.action('btn_undo', async (ctx) => {
    const serviceAccountAuth = getAuth();
    if (!serviceAccountAuth) return ctx.answerCbQuery('❌ service_account.json hilang.', { show_alert: true });
    
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = await getSheet(doc, serviceAccountAuth);
    if (!sheet) return ctx.answerCbQuery('❌ Gagal terhubung.', { show_alert: true });

    const rows = await sheet.getRows();
    if (rows.length === 0) return ctx.answerCbQuery('⚠️ Belum ada transaksi bulan ini.', { show_alert: true });

    const lastRow = rows[rows.length - 1];
    const infoTipe = lastRow.get('Tipe');
    const infoJumlah = lastRow.get('Jumlah');
    const infoKet = lastRow.get('Keterangan');

    try {
        await lastRow.delete();
        const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        await updateDashboard(doc, serviceAccountAuth, `${monthNames[new Date().getMonth()]} ${new Date().getFullYear()}`);
        ctx.answerCbQuery('✅ Transaksi Dibatalkan!', { show_alert: true });
        ctx.reply(`♻️ *Dibatalkan*:\n- ${infoTipe}\n- Rp ${infoJumlah}\n- ${infoKet}`, { parse_mode: 'Markdown' });
    } catch (err) {
        ctx.answerCbQuery('❌ Gagal menghapus transaksi.', { show_alert: true });
    }
});

bot.action('btn_ringkasan', async (ctx) => {
    ctx.answerCbQuery('Menghitung ringkasan & update Dashboard...');
    const serviceAccountAuth = getAuth();
    if (!serviceAccountAuth) return ctx.reply('❌ service_account.json hilang.');
    
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = await getSheet(doc, serviceAccountAuth);
    if (!sheet) return ctx.reply('❌ Gagal terhubung ke Google Sheets.');

    ctx.reply('✅ Ringkasan berhasil dihitung! Buka **Google Sheets** kamu untuk melihat Grafik lengkapnya.', { parse_mode: 'Markdown', ...mainMenu });
});

bot.action('btn_laporan', async (ctx) => {
    ctx.answerCbQuery('Menyiapkan dokumen...');
    const msg = await ctx.reply('⏳ Sedang di-generate...');
    
    const serviceAccountAuth = getAuth();
    if (!serviceAccountAuth) return ctx.telegram.editMessageText(ctx.chat.id, msg.message_id, null, '❌ service_account.json hilang.');
    
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = await getSheet(doc, serviceAccountAuth);
    const rows = await sheet.getRows();

    // 1. Generate CSV
    let csvStr = 'Tanggal,Tipe,Jumlah,Keterangan,Kategori\n';
    let totalIn = 0; let totalOut = 0;
    
    rows.forEach(r => {
        const j = parseFloat(r.get('Jumlah').replace(/[^\d]/g, ''));
        csvStr += `"${r.get('Tanggal')}","${r.get('Tipe')}","${j}","${r.get('Keterangan')}","${r.get('Kategori') || ''}"\n`;
        if (r.get('Tipe').includes('Pemasukan')) totalIn += j;
        else totalOut += j;
    });

    const csvBuffer = Buffer.from(csvStr, 'utf-8');

    // 2. Generate PDF
    const buffers = [];
    const pdfDoc = new PDFDocument({ margin: 50 });
    pdfDoc.on('data', buffers.push.bind(buffers));
    
    pdfDoc.on('end', async () => {
        const pdfData = Buffer.concat(buffers);
        await ctx.telegram.deleteMessage(ctx.chat.id, msg.message_id);
        
        await ctx.replyWithDocument({ source: pdfData, filename: `Laporan_${sheet.title}.pdf` }, { caption: '📄 Ini dia Laporan PDF kamu!' });
        await ctx.replyWithDocument({ source: csvBuffer, filename: `Data_${sheet.title}.csv` }, { caption: '📊 Ini dia file CSV (bisa dibuka di Excel) kamu!', ...mainMenu });
    });

    pdfDoc.fontSize(20).text(`Laporan Keuangan: ${sheet.title}`, { align: 'center' });
    pdfDoc.moveDown();
    pdfDoc.fontSize(12).text(`Total Pemasukan: Rp ${totalIn.toLocaleString('id-ID')}`, { align: 'left' });
    pdfDoc.text(`Total Pengeluaran: Rp ${totalOut.toLocaleString('id-ID')}`, { align: 'left' });
    pdfDoc.text(`Sisa Saldo: Rp ${(totalIn - totalOut).toLocaleString('id-ID')}`, { align: 'left' });
    pdfDoc.moveDown();
    pdfDoc.moveTo(50, pdfDoc.y).lineTo(550, pdfDoc.y).stroke();
    pdfDoc.moveDown();

    for(const r of rows) {
        const t = r.get('Tipe').includes('Pemasukan') ? '+' : '-';
        pdfDoc.fontSize(10).text(`[${r.get('Tanggal')}] ${r.get('Kategori') || 'Lain-lain'} | ${r.get('Keterangan')}`);
        pdfDoc.text(`${t} Rp ${r.get('Jumlah')} `, { align: 'right' });
        pdfDoc.moveDown(0.5);
    }
    
    pdfDoc.end();
});

bot.on('text', async (ctx) => {
    const text = ctx.message.text.trim();
    if (text.startsWith('/')) return;

    const match = text.match(/^(\+?)(\d+)\s+(.+)$/);

    if (match) {
        const isPemasukan = match[1] === '+';
        const jumlah = parseInt(match[2], 10);
        let keterangan = match[3].trim();
        const tipe = isPemasukan ? '📈 Pemasukan' : '📉 Pengeluaran';

        if (jumlah <= 0) return ctx.reply('⚠️ Jumlah uang harus lebih dari 0.');

        // Parse Reminder
        let reminderDate = null;
        const remMatch = keterangan.match(/\s+(?:tgl|tanggal|tiap)\s*(\d{1,2})$/i);
        if (remMatch) {
            const parsed = parseInt(remMatch[1], 10);
            if (parsed >= 1 && parsed <= 31) {
                reminderDate = parsed;
                keterangan = keterangan.replace(remMatch[0], '').trim();
            }
        }

        const kategori = isPemasukan ? 'Pemasukan' : getKategori(keterangan);

        const loadingMsg = await ctx.reply('⏳ Menyimpan...');
        const serviceAccountAuth = getAuth();
        if (!serviceAccountAuth) return ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ service_account.json hilang.');
        
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo();
        const sheet = await getSheet(doc, serviceAccountAuth);
        if (!sheet) return ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Gagal terhubung ke Google Sheets.');

        const dateStr = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });

        try {
            await sheet.addRow({ Tanggal: dateStr, Tipe: tipe, Jumlah: jumlah, Keterangan: keterangan, Kategori: kategori });
            
            // Set reminder if requested
            if (reminderDate) {
                await addReminder(doc, keterangan, reminderDate, ctx.chat.id);
            }

            const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
            await updateDashboard(doc, serviceAccountAuth, `${monthNames[new Date().getMonth()]} ${new Date().getFullYear()}`);

            let replyText = isPemasukan 
                ? `✅ *Pemasukan Dicatat!*\n\n📈 *Jumlah*: Rp ${jumlah.toLocaleString('id-ID')}\n📝 *Ket*: ${keterangan}`
                : `✅ *Pengeluaran Dicatat!*\n\n📉 *Jumlah*: Rp ${jumlah.toLocaleString('id-ID')}\n🏷️ *Kategori*: ${kategori}\n📝 *Ket*: ${keterangan}`;
            
            if (reminderDate) {
                replyText += `\n\n⏰ *Pengingat diaktifkan*: Bot akan mengingatkanmu tiap tanggal ${reminderDate}.`;
            }
            
            ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, replyText, { parse_mode: 'Markdown', ...mainMenu });
        } catch (error) {
            console.error("Error simpan data:", error);
            ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Terjadi kesalahan saat menyimpan data.');
        }
    } else {
        ctx.reply('❌ Format tidak dikenali.\n\nKetik `50000 kopi` untuk pengeluaran.\nKetik `+15000000 gaji` untuk pemasukan.', { parse_mode: 'Markdown' });
    }
});

bot.launch().then(() => {
    console.log("Bot Telegram sedang berjalan...");
});

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
