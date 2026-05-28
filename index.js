require('dotenv').config();
const { Telegraf, Markup } = require('telegraf');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

// 1. Inisialisasi Bot Telegram
const bot = new Telegraf(process.env.TELEGRAM_TOKEN);

// 2. Fungsi untuk mengakses Google Sheets
async function getSheet() {
    try {
        const creds = require('./service_account.json'); // Pastikan Anda membuat file ini!
        const serviceAccountAuth = new JWT({
            email: creds.client_email,
            key: creds.private_key,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
        await doc.loadInfo();
        const sheet = doc.sheetsByIndex[0];
        
        // Buat header jika kosong
        try {
            await sheet.setHeaderRow(['Tanggal', 'Tipe', 'Jumlah', 'Keterangan']);
        } catch (e) {
            // Abaikan jika header sudah ada
        }
        return sheet;
    } catch (error) {
        console.error("Gagal terhubung ke Google Sheets:", error);
        return null;
    }
}

// 3. Command /start & Menu Utama
bot.start((ctx) => {
    ctx.reply(
        'Halo! 👋 Saya adalah Bot Pencatat Keuangan Anda.\n\n' +
        '🟢 **Pemasukan:** Ketik `+<angka> <keterangan>`\nContoh: `+15000000 gaji`\n\n' +
        '🔴 **Pengeluaran:** Ketik `<angka> <keterangan>`\nContoh: `50000 kopi`',
        Markup.keyboard([
            ['📊 Ringkasan Keuangan'],
            ['💡 Bantuan']
        ]).resize()
    );
});

// 4. Tombol Bantuan
bot.hears('💡 Bantuan', (ctx) => {
    ctx.reply('Cara pakai sangat mudah:\n- Untuk pengeluaran, langsung ketik angkanya, contoh: `50000 makan siang`\n- Untuk pemasukan, gunakan tanda plus, contoh: `+2000000 bonus`');
});

// 5. Tombol Ringkasan
bot.hears('📊 Ringkasan Keuangan', async (ctx) => {
    ctx.reply('⏳ Sedang menghitung ringkasan dari Spreadsheet...');
    const sheet = await getSheet();
    if (!sheet) return ctx.reply('❌ Gagal terhubung ke Google Sheets. Pastikan service_account.json sudah ada.');

    const rows = await sheet.getRows();
    let totalPemasukan = 0;
    let totalPengeluaran = 0;

    rows.forEach(row => {
        const tipe = row.get('Tipe');
        const jumlah = parseFloat(row.get('Jumlah')) || 0;
        if (tipe === 'pemasukan') totalPemasukan += jumlah;
        if (tipe === 'pengeluaran') totalPengeluaran += jumlah;
    });

    const saldo = totalPemasukan - totalPengeluaran;
    const pesan = `📊 **Ringkasan Keuangan**\n\n🟢 Total Pemasukan: Rp ${totalPemasukan.toLocaleString('id-ID')}\n🔴 Total Pengeluaran: Rp ${totalPengeluaran.toLocaleString('id-ID')}\n\n💰 **Saldo Saat Ini: Rp ${saldo.toLocaleString('id-ID')}**`;
    ctx.reply(pesan);
});

// 6. Logika Parsing Pesan (Pemasukan & Pengeluaran)
bot.on('text', async (ctx) => {
    const text = ctx.message.text.trim();
    
    // Jangan proses jika text adalah tombol menu
    if (text === '📊 Ringkasan Keuangan' || text === '💡 Bantuan') return;

    // Regex untuk mendeteksi pesan: 
    // ^(\+?) -> opsional tanda plus (Group 1)
    // (\d+) -> angka jumlah uang (Group 2)
    // \s+(.+) -> spasi lalu keterangan (Group 3)
    const match = text.match(/^(\+?)(\d+)\s+(.+)$/);

    if (match) {
        const isPemasukan = match[1] === '+';
        const jumlah = parseInt(match[2], 10);
        const keterangan = match[3].trim();
        const tipe = isPemasukan ? 'pemasukan' : 'pengeluaran';

        const loadingMsg = await ctx.reply('⏳ Sedang menyimpan ke Spreadsheet...');

        const sheet = await getSheet();
        if (!sheet) {
            return ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Gagal terhubung ke Google Sheets. Pastikan service_account.json sudah diatur.');
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
                ? `✅ **Pemasukan Berhasil Dicatat!**\n💰 Jumlah: Rp ${jumlah.toLocaleString('id-ID')}\n📝 Keterangan: ${keterangan}`
                : `✅ **Pengeluaran Berhasil Dicatat!**\n💸 Jumlah: Rp ${jumlah.toLocaleString('id-ID')}\n📝 Keterangan: ${keterangan}`;
            
            ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, replyText);
        } catch (error) {
            ctx.telegram.editMessageText(ctx.chat.id, loadingMsg.message_id, null, '❌ Terjadi kesalahan saat menyimpan data.');
        }

    } else {
        ctx.reply('❌ Format tidak dikenali. \nGunakan format:\n`50000 kopi` (Pengeluaran)\n`+15000000 gaji` (Pemasukan)');
    }
});

bot.launch().then(() => {
    console.log("Bot Telegram sedang berjalan...");
});

// Enable graceful stop
process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
