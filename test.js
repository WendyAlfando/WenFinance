require('dotenv').config();
const { GoogleGenerativeAI } = require('@google/generative-ai');
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

async function test() {
    try {
        const text = "tadi gw beli risol 10500";
        const model = genAI.getGenerativeModel({ model: "gemini-flash-latest" });
        const prompt = `Ekstrak satu atau beberapa transaksi keuangan dari teks ini: "${text}". Output HANYA JSON array: [{"tipe": "Pengeluaran"|"Pemasukan", "jumlah": <angka>, "keterangan": "<deskripsi>"}] tanpa backtick atau markdown tambahan.`;
        const result = await model.generateContent(prompt);
        console.log("Raw Response:", result.response.text());
        
        let jsonText = result.response.text().replace(/```json/g, '').replace(/```/g, '').trim();
        console.log("Cleaned JSON:", jsonText);
        
        const parsed = JSON.parse(jsonText);
        console.log("Parsed:", parsed);
    } catch (e) {
        console.error("ERROR", e);
    }
}
test();
