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

// 3. Helper Bulan Ini (DD/MM/YYYY)
function isCurrentMonth(dateStr) {
    if (!dateStr) return false;
    const parts = dateStr.split(/[\/, ]/);
    if (parts.length >= 2) {
        const d = new Date();
        return parseInt(parts[1]) === (d.getMonth() + 1) && parseInt(parts[2]) === d.getFullYear();
    }
    return false;
}

// 4. Menu Telegram
const mainMenu = Markup.inlineKeyboard([
    [Markup.button.callback('📊 Ringkasan Bulanan', 'btn_ringkasan'), Markup.button.callback('📈 Grafik Harian', 'btn_grafik_harian')],
    [Markup.button.callback('📄 Export Laporan PDF', 'btn_laporan')],
    [Markup.button.url('🌐 Buka Dashboard Visual', 'http://146.190.85.119:3000')]
]);

bot.start(async (ctx) => {
    await ctx.reply('Memperbarui sistem menu...', Markup.removeKeyboard());
    ctx.reply(
        'Halo! 👋 Saya adalah Bot Pencatat Keuangan Cerdas.\n\n' +
        '🤖 *AI Mode:* Ketik santai "tadi siang makan ayam 50rb pakai qris"!\n' +
        'Gunakan tombol di bawah 👇',
        { parse_mode: 'Markdown', ...mainMenu }
    );
});

bot.command('menu', async (ctx) => {
    ctx.reply('🎛 *Menu Utama*', { parse_mode: 'Markdown', ...mainMenu });
});

bot.action('btn_ringkasan', async (ctx) => {
    ctx.answerCbQuery('Menghitung ringkasan & membuat grafik...');
    const auth = getAuth();
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    await doc.loadInfo();
    
    const sheet = doc.sheetsByTitle['Pengeluaran'];
    if (!sheet) return ctx.reply('❌ Tab Pengeluaran tidak ditemukan.');
    const rows = await sheet.getRows();
    
    let catTotals = {};
    for (const row of rows) {
        if (isCurrentMonth(row.get('Tanggal'))) {
            const cat = row.get('Kategori') || 'Lain-lain';
            const val = parseFloat((row.get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, ''));
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

bot.action('btn_grafik_harian', async (ctx) => {
    ctx.answerCbQuery('Membuat grafik harian...');
    const auth = getAuth();
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
    await doc.loadInfo();
    
    const sheet = doc.sheetsByTitle['Pengeluaran'];
    if (!sheet) return ctx.reply('❌ Tab Pengeluaran tidak ditemukan.');
    const rows = await sheet.getRows();
    
    let dailyTotals = {};
    for (const row of rows) {
        const dateStr = row.get('Tanggal') || '';
        if (isCurrentMonth(dateStr)) {
            const dayMatch = dateStr.match(/^(\d{2})\//);
            if (dayMatch) {
                const day = dayMatch[1];
                const val = parseFloat((row.get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, ''));
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
        
        const sheetIn = doc.sheetsByTitle['Pendapatan'];
        const sheetOut = doc.sheetsByTitle['Pengeluaran'];
        
        let totalPemasukan = 0;
        let totalPengeluaran = 0;
        const transactions = [];

        if (sheetIn) {
            const rows = await sheetIn.getRows();
            for (const row of rows) {
                if (isCurrentMonth(row.get('Tanggal'))) {
                    const jumlah = parseFloat((row.get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, ''));
                    totalPemasukan += jumlah;
                    transactions.push({ tgl: row.get('Tanggal').split(',')[0], tipe: '+', jumlah: jumlah, ket: row.get('Keterangan') });
                }
            }
        }
        
        if (sheetOut) {
            const rows = await sheetOut.getRows();
            for (const row of rows) {
                if (isCurrentMonth(row.get('Tanggal'))) {
                    const jumlah = parseFloat((row.get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, ''));
                    totalPengeluaran += jumlah;
                    transactions.push({ tgl: row.get('Tanggal').split(',')[0], tipe: '-', jumlah: jumlah, ket: row.get('Deskripsi') });
                }
            }
        }
        
        transactions.sort((a,b) => parseInt(a.tgl.split('/')[0]) - parseInt(b.tgl.split('/')[0]));

        const pdf = new PDFDocument({ margin: 50 });
        const buffers = [];
        pdf.on('data', buffers.push.bind(buffers));
        
        pdf.fontSize(20).text('Laporan Keuangan Bulanan', { align: 'center' });
        pdf.moveDown();
        const currentMonthTitle = `${["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][new Date().getMonth()]} ${new Date().getFullYear()}`;
        pdf.fontSize(14).text(`Bulan: ${currentMonthTitle}`);
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
            await ctx.replyWithDocument({ source: pdfData, filename: `Laporan_${currentMonthTitle.replace(' ', '_')}.pdf` });
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

        if (!genAI) return await ctx.reply('❌ GEMINI API KEY tidak ada.', { parse_mode: 'Markdown' });
        const loadingAI = await ctx.reply('🤖 Memproses dengan AI...');
        
        let transactions = [];
        try {
            const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
            const prompt = `Ekstrak transaksi keuangan dari teks ini: "${text}". 
Output HANYA array JSON murni tanpa markdown, format: 
[
  {
    "tipe": "Pengeluaran" atau "Pemasukan",
    "jumlah": <angka tanpa Rp/titik>,
    "keterangan": "deskripsi barang/sumber",
    "kategori": "Pilih Kategori: Makan & Minum, Transportasi, Belanja, Hiburan, Tagihan & Utilitas, Kesehatan, Pendidikan, Cicilan, Internet & Pulsa, Kecantikan, Asuransi, Lainnya (Jika pemasukan: Gaji, Freelance, Bonus, Investasi, Bisnis, Lainnya)",
    "metode_bayar": "Transfer Bank, QRIS, Cash, Debit BCA, GoPay, OVO, ShopeePay, Dana, dll",
    "sumber": "Dari mana asalnya (hanya jika Pemasukan, misal 'Gaji Bulanan', 'Proyek X')",
    "kebutuhan": "Ya" atau "Tidak" (hanya jika Pengeluaran, apakah ini kebutuhan atau keinginan?)
  }
]`;
            const result = await model.generateContent(prompt);
            let rawText = result.response.text();
            let start = rawText.indexOf('[');
            let end = rawText.lastIndexOf(']');
            if (start === -1 || end === -1) throw new Error("Tidak ada JSON");
            
            const parsed = JSON.parse(rawText.substring(start, end + 1));
            transactions = Array.isArray(parsed) ? parsed : [parsed];
            await ctx.telegram.deleteMessage(ctx.chat.id, loadingAI.message_id).catch(() => {});
        } catch (e) {
            return await ctx.telegram.editMessageText(ctx.chat.id, loadingAI.message_id, null, `❌ AI Error: ${e.message}`).catch(() => {});
        }

        if (transactions.length === 0) return;

        const loadingMsg = await ctx.reply('⏳ Menyimpan ke Spreadsheet...');
        try {
            const auth = getAuth();
            const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
            await doc.loadInfo();
            const dateStr = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });

            let finalReply = "✅ *Transaksi Berhasil Dicatat!*\n";
            for (const tx of transactions) {
                if (tx.tipe === 'Pemasukan') {
                    let sheet = doc.sheetsByTitle['Pendapatan'];
                    if (!sheet) throw new Error("Tab Pendapatan tidak ditemukan!");
                    // ['Tanggal', 'Sumber', 'Kategori', 'Jumlah (Rp)', 'Metode', 'Keterangan']
                    await sheet.addRow({ 
                        'Tanggal': dateStr, 
                        'Sumber': tx.sumber || tx.keterangan, 
                        'Kategori': tx.kategori || 'Lainnya', 
                        'Jumlah (Rp)': tx.jumlah, 
                        'Metode': tx.metode_bayar || 'Cash', 
                        'Keterangan': tx.keterangan 
                    });
                } else {
                    let sheet = doc.sheetsByTitle['Pengeluaran'];
                    if (!sheet) throw new Error("Tab Pengeluaran tidak ditemukan!");
                    // ['Tanggal', 'Kategori', 'Deskripsi', 'Jumlah (Rp)', 'Metode Bayar', 'Kebutuhan?', 'Catatan']
                    await sheet.addRow({ 
                        'Tanggal': dateStr, 
                        'Kategori': tx.kategori || 'Lainnya', 
                        'Deskripsi': tx.keterangan, 
                        'Jumlah (Rp)': tx.jumlah, 
                        'Metode Bayar': tx.metode_bayar || 'Cash', 
                        'Kebutuhan?': tx.kebutuhan || 'Ya', 
                        'Catatan': '' 
                    });
                }

                const icon = tx.tipe === 'Pemasukan' ? '📈' : '📉';
                finalReply += `\n${icon} Rp ${tx.jumlah.toLocaleString('id-ID')} | ${tx.keterangan} (${tx.metode_bayar || 'Cash'})`;
            }

            await ctx.telegram.deleteMessage(ctx.chat.id, loadingMsg.message_id).catch(() => {});
            await ctx.reply(finalReply, { parse_mode: 'Markdown', ...mainMenu }).catch(() => {});
        } catch (err) {
            await ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, `❌ Gagal menyimpan: ${err.message}`).catch(() => {});
        }
    } catch (e) {}
});

bot.on('photo', async (ctx) => {
    if (!genAI) return ctx.reply('❌ GEMINI API KEY tidak ada.');
    const loading = await ctx.reply('📸 Membaca struk...');
    try {
        const photo = ctx.message.photo[ctx.message.photo.length - 1];
        const fileLink = await ctx.telegram.getFileLink(photo.file_id);
        const response = await axios.get(fileLink.href, { responseType: 'arraybuffer' });
        const base64Data = Buffer.from(response.data, 'binary').toString('base64');
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const prompt = `Ini foto struk. Output HANYA JSON murni: {"jumlah": <angka_total>, "keterangan": "Belanja di <Toko>", "kategori": "Kategori yang cocok (Makan & Minum, Belanja, Transportasi, dll)", "metode_bayar": "Cash/Debit/QRIS (tebak)", "kebutuhan": "Ya/Tidak"}`;
        const result = await model.generateContent([{ inlineData: { data: base64Data, mimeType: 'image/jpeg' } }, prompt]);
        let rawText = result.response.text();
        let start = rawText.indexOf('{');
        let end = rawText.lastIndexOf('}');
        const parsed = JSON.parse(rawText.substring(start, end + 1));
        
        const auth = getAuth();
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['Pengeluaran'];
        const dateStr = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
        
        await sheet.addRow({ 
            'Tanggal': dateStr, 
            'Kategori': parsed.kategori || 'Belanja', 
            'Deskripsi': parsed.keterangan, 
            'Jumlah (Rp)': parsed.jumlah, 
            'Metode Bayar': parsed.metode_bayar || 'Cash', 
            'Kebutuhan?': parsed.kebutuhan || 'Ya', 
            'Catatan': 'Via Struk' 
        });
        
        await ctx.telegram.deleteMessage(ctx.chat.id, loading.message_id).catch(()=>{});
        ctx.reply(`✅ **Struk Dicatat!**\n📉 Rp ${parseInt(parsed.jumlah).toLocaleString('id-ID')} | ${parsed.keterangan} (${parsed.metode_bayar})`, {parse_mode: 'Markdown'});
    } catch(e) { await ctx.telegram.editMessageText(ctx.chat.id, loading.message_id, null, `❌ Gagal: ${e.message}`); }
});

app.get('/api/dashboard', async (req, res) => {
    try {
        const auth = getAuth();
        if (!auth) return res.status(500).json({error: 'Auth failed'});
        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, auth);
        await doc.loadInfo();
        
        const sheetIn = doc.sheetsByTitle['Pendapatan'];
        const sheetOut = doc.sheetsByTitle['Pengeluaran'];
        
        let totalMasuk = 0; let totalKeluar = 0; let dailyTotals = {}; let transactions = [];
        let expenseCatTotals = {}; let incomeCatTotals = {};
        let wallets = {};

        if (sheetIn) {
            const rows = await sheetIn.getRows();
            for (let i = Math.max(0, rows.length - 15); i < rows.length; i++) {
                if (isCurrentMonth(rows[i].get('Tanggal'))) {
                    transactions.push({ tgl: rows[i].get('Tanggal').split(',')[0], ket: rows[i].get('Keterangan'), kat: rows[i].get('Kategori'), tipe: '📈 Pemasukan', jumlah: parseFloat((rows[i].get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, '')) });
                }
            }
            for (const row of rows) {
                if (isCurrentMonth(row.get('Tanggal'))) {
                    const jumlah = parseFloat((row.get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, ''));
                    totalMasuk += jumlah;
                    incomeCatTotals[row.get('Kategori') || 'Pemasukan'] = (incomeCatTotals[row.get('Kategori') || 'Pemasukan'] || 0) + jumlah;
                    
                    const dompet = row.get('Metode') || 'Cash';
                    wallets[dompet] = (wallets[dompet] || 0) + jumlah;
                }
            }
        }

        if (sheetOut) {
            const rows = await sheetOut.getRows();
            for (let i = Math.max(0, rows.length - 15); i < rows.length; i++) {
                if (isCurrentMonth(rows[i].get('Tanggal'))) {
                    transactions.push({ tgl: rows[i].get('Tanggal').split(',')[0], ket: rows[i].get('Deskripsi'), kat: rows[i].get('Kategori'), tipe: '📉 Pengeluaran', jumlah: parseFloat((rows[i].get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, '')) });
                }
            }
            for (const row of rows) {
                const dateStr = row.get('Tanggal') || '';
                if (isCurrentMonth(dateStr)) {
                    const jumlah = parseFloat((row.get('Jumlah (Rp)') || '0').replace(/[^\d.-]/g, ''));
                    totalKeluar += jumlah;
                    expenseCatTotals[row.get('Kategori') || 'Lain-lain'] = (expenseCatTotals[row.get('Kategori') || 'Lain-lain'] || 0) + jumlah;
                    
                    const dayMatch = dateStr.match(/^(\d{2})\//);
                    if (dayMatch) dailyTotals[dayMatch[1]] = (dailyTotals[dayMatch[1]] || 0) + jumlah;

                    const dompet = row.get('Metode Bayar') || 'Cash';
                    wallets[dompet] = (wallets[dompet] || 0) - jumlah;
                }
            }
        }
        
        transactions.sort((a,b) => parseInt(b.tgl.split('/')[0]) - parseInt(a.tgl.split('/')[0]));
        transactions = transactions.slice(0, 15); // Ambil 15 terakhir

        const expenseCategories = Object.keys(expenseCatTotals).map(k => ({ name: k, amount: expenseCatTotals[k] })).sort((a,b)=>b.amount-a.amount);
        const incomeCategories = Object.keys(incomeCatTotals).map(k => ({ name: k, amount: incomeCatTotals[k] })).sort((a,b)=>b.amount-a.amount);
        const dailyLabels = Object.keys(dailyTotals).sort((a,b)=>parseInt(a)-parseInt(b));
        const dailyData = dailyLabels.map(l => dailyTotals[l]);
        
        const currentMonthTitle = `${["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][new Date().getMonth()]} ${new Date().getFullYear()}`;
        res.json({ month: currentMonthTitle, totalMasuk, totalKeluar, expenseCategories, incomeCategories, dailyLabels, dailyData, transactions, wallets });
    } catch(e) { res.status(500).json({error: e.message}); }
});

app.listen(3000, () => console.log('Web App berjalan di port 3000'));
bot.launch().then(() => console.log("Bot Telegram sedang berjalan..."));
process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
