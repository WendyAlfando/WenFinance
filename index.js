require('dotenv').config();
const { Telegraf, Markup } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');

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

// 3. Fungsi Update Dashboard Google Sheets
async function updateDashboard(doc, serviceAccountAuth, currentMonthTitle) {
    let allTimeMasuk = 0;
    let allTimeKeluar = 0;
    let monthMasuk = 0;
    let monthKeluar = 0;

    for (const sheet of doc.sheetsByIndex) {
        if (sheet.title === '📊 Dashboard Utama') continue;
        
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
                else monthKeluar += jumlah;
            }
        }
    }

    let dashboardSheet = doc.sheetsByTitle['📊 Dashboard Utama'];
    let isNew = false;
    
    if (!dashboardSheet) {
        dashboardSheet = await doc.addSheet({
            title: '📊 Dashboard Utama',
            index: 0,
            gridProperties: { columnCount: 10, rowCount: 20 }
        });
        isNew = true;
    } else {
        if (dashboardSheet.index !== 0) {
            await dashboardSheet.updateProperties({ index: 0 });
        }
    }

    await dashboardSheet.loadCells('A1:C12');

    // Title 1
    const t1 = dashboardSheet.getCell(0, 0); // A1
    t1.value = `Ringkasan Bulan Ini (${currentMonthTitle})`;
    t1.textFormat = { bold: true, fontSize: 12 };
    
    dashboardSheet.getCell(1, 0).value = "Bulan";
    dashboardSheet.getCell(1, 1).value = currentMonthTitle;
    
    dashboardSheet.getCell(2, 0).value = "Pemasukan";
    dashboardSheet.getCell(2, 1).value = monthMasuk;
    
    dashboardSheet.getCell(3, 0).value = "Pengeluaran";
    dashboardSheet.getCell(3, 1).value = monthKeluar;

    dashboardSheet.getCell(4, 0).value = "Saldo";
    dashboardSheet.getCell(4, 1).value = monthMasuk - monthKeluar;

    // Title 2
    const t2 = dashboardSheet.getCell(6, 0); // A7
    t2.value = "Ringkasan Semua Bulan";
    t2.textFormat = { bold: true, fontSize: 12 };

    dashboardSheet.getCell(7, 0).value = "Total Pemasukan";
    dashboardSheet.getCell(7, 1).value = allTimeMasuk;

    dashboardSheet.getCell(8, 0).value = "Total Pengeluaran";
    dashboardSheet.getCell(8, 1).value = allTimeKeluar;

    dashboardSheet.getCell(9, 0).value = "Sisa Saldo";
    dashboardSheet.getCell(9, 1).value = allTimeMasuk - allTimeKeluar;

    await dashboardSheet.saveUpdatedCells();

    if (isNew) {
        // Add charts & formatting using raw request
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
                                        title: "Distribusi Bulan Ini",
                                        pieChart: {
                                            legendPosition: "RIGHT_LEGEND",
                                            domain: { sourceRange: { sources: [{ sheetId, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 1 }] } },
                                            series: { sourceRange: { sources: [{ sheetId, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 1, endColumnIndex: 2 }] } }
                                        }
                                    },
                                    position: { overlayPosition: { anchorCell: { sheetId, rowIndex: 0, columnIndex: 3 }, widthPixels: 400, heightPixels: 280 } }
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
                                    position: { overlayPosition: { anchorCell: { sheetId, rowIndex: 6, columnIndex: 3 }, widthPixels: 400, heightPixels: 280 } }
                                }
                            }
                        }
                    ]
                }
            });
        } catch (err) {
            console.error("Gagal menambahkan chart ke dashboard:", err);
        }
    }
}

// 4. Fungsi untuk mengakses/membuat sheet transaksi bulan ini
async function getSheet(doc, serviceAccountAuth) {
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const date = new Date();
    const currentMonthTitle = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
    
    let sheet = doc.sheetsByTitle[currentMonthTitle];
    
    if (!sheet) {
        sheet = await doc.addSheet({
            title: currentMonthTitle,
            headerValues: ['Tanggal', 'Tipe', 'Jumlah', 'Keterangan'],
            gridProperties: { frozenRowCount: 1 }
        });

        await sheet.loadCells('A1:D1');
        for (let i = 0; i < 4; i++) {
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
    }
    
    // Perbarui dashboard setiap kali sheet bulan ini diakses (artinya ada interaksi)
    await updateDashboard(doc, serviceAccountAuth, currentMonthTitle);
    
    return sheet;
}

// 5. Command /start & Menu Utama Telegram
const mainMenu = Markup.inlineKeyboard([
    [Markup.button.callback('📊 Ringkasan & Update Dashboard', 'btn_ringkasan')],
    [Markup.button.callback('🔙 Batal Terakhir', 'btn_undo'), Markup.button.callback('💡 Bantuan', 'btn_help')]
]);

bot.start(async (ctx) => {
    await ctx.reply('Memperbarui sistem menu...', Markup.removeKeyboard());
    ctx.reply(
        'Halo! 👋 Saya adalah Bot Pencatat Keuangan Anda.\n\n' +
        '🟢 *Pemasukan:* Ketik `+<angka> <keterangan>`\nContoh: `+15000000 gaji bulan ini`\n\n' +
        '🔴 *Pengeluaran:* Ketik `<angka> <keterangan>`\nContoh: `50000 makan siang`\n\n' +
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
        '- *Pengeluaran*: Langsung ketik angkanya.\n  Contoh: `50000 bensin`\n' +
        '- *Pemasukan*: Gunakan tanda plus `+`.\n  Contoh: `+2000000 bonus project`\n' +
        '- *Hapus Data Terakhir*: Klik tombol Batal Terakhir.\n' +
        '- *Laporan*: Buka file Google Sheets-mu untuk melihat grafik otomatis di tab **📊 Dashboard Utama**.',
        { parse_mode: 'Markdown' }
    );
});

bot.action('btn_undo', async (ctx) => {
    const serviceAccountAuth = getAuth();
    if (!serviceAccountAuth) return ctx.answerCbQuery('❌ service_account.json hilang.', { show_alert: true });
    
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    
    const sheet = await getSheet(doc, serviceAccountAuth);
    if (!sheet) return ctx.answerCbQuery('❌ Gagal terhubung ke Sheets.', { show_alert: true });

    const rows = await sheet.getRows();
    if (rows.length === 0) {
        return ctx.answerCbQuery('⚠️ Belum ada transaksi bulan ini.', { show_alert: true });
    }

    const lastRow = rows[rows.length - 1];
    const infoTipe = lastRow.get('Tipe');
    const infoJumlah = lastRow.get('Jumlah');
    const infoKet = lastRow.get('Keterangan');

    try {
        await lastRow.delete();
        // Update dashboard setelah hapus
        const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        await updateDashboard(doc, serviceAccountAuth, `${monthNames[new Date().getMonth()]} ${new Date().getFullYear()}`);
        
        ctx.answerCbQuery('✅ Transaksi Berhasil Dibatalkan!', { show_alert: true });
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

    ctx.reply('✅ Ringkasan bulan ini berhasil dihitung!\nSilakan buka **Google Sheets** kamu untuk melihat Grafik di tab **📊 Dashboard Utama**.', { parse_mode: 'Markdown', ...mainMenu });
});

bot.on('text', async (ctx) => {
    const text = ctx.message.text.trim();
    if (text.startsWith('/')) return;

    const match = text.match(/^(\+?)(\d+)\s+(.+)$/);

    if (match) {
        const isPemasukan = match[1] === '+';
        const jumlah = parseInt(match[2], 10);
        const keterangan = match[3].trim();
        const tipe = isPemasukan ? '📈 Pemasukan' : '📉 Pengeluaran';

        if (jumlah <= 0) {
            return ctx.reply('⚠️ Jumlah uang harus lebih dari 0.');
        }

        const loadingMsg = await ctx.reply('⏳ Menyimpan...');

        const serviceAccountAuth = getAuth();
        if (!serviceAccountAuth) return ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ service_account.json hilang.');
        
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo();

        const sheet = await getSheet(doc, serviceAccountAuth);
        if (!sheet) {
            return ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Gagal terhubung ke Google Sheets.');
        }

        const dateOptions = { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' };
        const dateStr = new Date().toLocaleString('id-ID', dateOptions);

        try {
            await sheet.addRow({
                Tanggal: dateStr,
                Tipe: tipe,
                Jumlah: jumlah,
                Keterangan: keterangan
            });
            
            // Update dashboard setelah tambah baris
            const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
            await updateDashboard(doc, serviceAccountAuth, `${monthNames[new Date().getMonth()]} ${new Date().getFullYear()}`);

            const replyText = isPemasukan 
                ? `✅ *Pemasukan Dicatat!*\n\n📈 *Jumlah*: Rp ${jumlah.toLocaleString('id-ID')}\n📝 *Ket*: ${keterangan}`
                : `✅ *Pengeluaran Dicatat!*\n\n📉 *Jumlah*: Rp ${jumlah.toLocaleString('id-ID')}\n📝 *Ket*: ${keterangan}`;
            
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
