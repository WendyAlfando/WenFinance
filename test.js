const { GoogleGenerativeAI } = require('@google/generative-ai');
const genAI = new GoogleGenerativeAI('AIzaSyBfmKlKh6Qgo1TffPePxgLQSIA-eFnKz2M');

async function test() {
    try {
        const model = genAI.getGenerativeModel({ model: "gemini-flash-latest" });
        const result = await model.generateContent("halo");
        console.log(result.response.text());
    } catch (e) {
        console.error("ERROR", e);
    }
}
test();
