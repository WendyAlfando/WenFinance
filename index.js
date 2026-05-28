require('dotenv').config();
const { Telegraf, Markup } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');
const express = require('express');
const cors = require('cors');

// 1. Inisialisasi Express & API
const app = express();
app.use(cors());

// 2. Inisialisasi Bot Telegram
const bot = new Telegraf(process.env.TELEGRAM_TOKEN);

// 3. Fungsi untuk mengakses Google Sheets
async function getSheet() {
    try {
        if (!fs.existsSync('./service_account.json')) {
            console.error("service_account.json tidak ditemukan!");
            return null;
        }
        const creds = require('./service_account.json');
        const serviceAccountAuth = new JWT({
            email: creds.client_email,
            key: creds.private_key,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo();
        
        const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        const date = new Date();
        const sheetTitle = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
        
        let sheet = doc.sheetsByTitle[sheetTitle];
        
        if (!sheet) {
            sheet = await doc.addSheet({
                title: sheetTitle,
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
            } catch (err) {
                console.error("Gagal melakukan batchUpdate format:", err);
            }
        }
        
        return sheet;
    } catch (error) {
        console.error("Gagal terhubung ke Google Sheets:", error);
        return null;
    }
}

// 4. API Endpoints untuk Dashboard Web
app.get('/api/data', async (req, res) => {
    try {
        const sheet = await getSheet();
        if (!sheet) return res.status(500).json({ error: 'Gagal terhubung ke Sheets' });

        const rows = await sheet.getRows();
        let totalPemasukan = 0;
        let totalPengeluaran = 0;
        let transactions = [];

        rows.forEach(row => {
            const tipe = row.get('Tipe');
            const jumlahStr = row.get('Jumlah') || '0';
            const jumlah = parseFloat(jumlahStr.toString().replace(/[^\d.-]/g, ''));
            const tanggal = row.get('Tanggal');
            const keterangan = row.get('Keterangan');
            
            const isIncome = tipe === '📈 Pemasukan' || tipe === 'pemasukan';
            if (isIncome) totalPemasukan += jumlah;
            else totalPengeluaran += jumlah;

            transactions.push({
                id: row.rowNumber,
                tanggal,
                tipe: isIncome ? 'Pemasukan' : 'Pengeluaran',
                jumlah,
                keterangan
            });
        });

        // 10 transaksi terakhir
        const recentTransactions = transactions.reverse().slice(0, 10);

        res.json({
            month: sheet.title,
            saldo: totalPemasukan - totalPengeluaran,
            totalPemasukan,
            totalPengeluaran,
            recentTransactions
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 5. Command /start & Menu Utama Telegram
const mainMenu = Markup.inlineKeyboard([
    [Markup.button.callback('📊 Ringkasan Bulan Ini', 'btn_ringkasan')],
    [Markup.button.url('🌐 Buka Dashboard Web', 'http://localhost:5173')],
    [Markup.button.callback('🔙 Batal Terakhir', 'btn_undo'), Markup.button.callback('💡 Bantuan', 'btn_help')]
]);

bot.start((ctx) => {
    ctx.reply(
        'Halo! 👋 Saya adalah Bot Pencatat Keuangan Anda.\n\n' +
        '🟢 *Pemasukan:* Ketik `+<angka> <keterangan>`\nContoh: `+15000000 gaji bulan ini`\n\n' +
        '🔴 *Pengeluaran:* Ketik `<angka> <keterangan>`\nContoh: `50000 makan siang`\n\n' +
        'Gunakan tombol di bawah untuk mengakses menu 👇',
        { parse_mode: 'Markdown', ...mainMenu }
    );
});

bot.command('menu', (ctx) => {
    ctx.reply('🎛 *Menu Utama*', { parse_mode: 'Markdown', ...mainMenu });
});

bot.action('btn_help', (ctx) => {
    ctx.answerCbQuery();
    ctx.reply(
        '📘 *Cara Menggunakan Bot:*\n\n' +
        '- *Pengeluaran*: Langsung ketik angkanya.\n  Contoh: `50000 bensin`\n' +
        '- *Pemasukan*: Gunakan tanda plus `+`.\n  Contoh: `+2000000 bonus project`\n' +
        '- *Hapus Data Terakhir*: Klik tombol Batal Terakhir.\n' +
        '- *Laporan*: Data disimpan otomatis tiap bulan dalam tab baru (misal: "Mei 2026").',
        { parse_mode: 'Markdown' }
    );
});

bot.action('btn_undo', async (ctx) => {
    const sheet = await getSheet();
    if (!sheet) return ctx.answerCbQuery('❌ Gagal terhubung ke Google Sheets.', { show_alert: true });

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
        ctx.answerCbQuery('✅ Transaksi Berhasil Dibatalkan!', { show_alert: true });
        ctx.reply(`♻️ *Dibatalkan*:\n- ${infoTipe}\n- Rp ${infoJumlah}\n- ${infoKet}`, { parse_mode: 'Markdown' });
    } catch (err) {
        ctx.answerCbQuery('❌ Gagal menghapus transaksi.', { show_alert: true });
    }
});

bot.action('btn_ringkasan', async (ctx) => {
    ctx.answerCbQuery('Menghitung ringkasan...');
    const sheet = await getSheet();
    if (!sheet) return ctx.reply('❌ Gagal terhubung ke Google Sheets.');

    const rows = await sheet.getRows();
    let totalPemasukan = 0;
    let totalPengeluaran = 0;

    rows.forEach(row => {
        const tipe = row.get('Tipe');
        const jumlahStr = row.get('Jumlah') || '0';
        const jumlah = parseFloat(jumlahStr.toString().replace(/[^\d.-]/g, ''));
        
        if (tipe === '📈 Pemasukan' || tipe === 'pemasukan') totalPemasukan += jumlah;
        if (tipe === '📉 Pengeluaran' || tipe === 'pengeluaran') totalPengeluaran +=াস্থ্য totalPengeluaran += jumlah;
    });

    const saldo = totalPemasukan - totalPengeluaran;
    const pesan = `📊 *Ringkasan Keuangan (${sheet.title})*\n\n` +
                  `🟢 *Pemasukan*: Rp ${totalPemasukan.toLocaleString('id-ID')}\n` +
                  `🔴 *Pengeluaran*: Rp ${totalPengeluaran.toLocaleString('id-ID')}\n\n` +
                  `💰 *Saldo Saat Ini: Rp ${saldo.toLocaleString('id-ID')}*`;
                  
    ctx.reply(pesan, { parse_mode: 'Markdown', ...mainMenu });
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

        const sheet = await getSheet();
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

app.listen(3000, () => console.log('API Dashboard berjalan di port 3000'));

bot.launch().then(() => {
    console.log("Bot Telegram sedang berjalan...");
});

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
