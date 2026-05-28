async function test() {
    try {
        const models = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=AIzaSyBfmKlKh6Qgo1TffPePxgLQSIA-eFnKz2M`).then(r => r.json());
        console.log(models.models.filter(m => m.name.includes('flash') && m.supportedGenerationMethods.includes('generateContent')).map(m => m.name));
    } catch (e) {
        console.error(e);
    }
}
test();
