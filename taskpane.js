// taskpane.js
(() => {
  Office.onReady(() => {
    document.getElementById('checkBtn').addEventListener('click', checkText);
  });

  async function checkText() {
    try {
      await Word.run(async context => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();

        const text = selection.text.trim();
        if (!text) {
          document.getElementById('results').textContent = 'Selecciona texto primero.';
          return;
        }

        const prompt = `Corrige ortografía y gramática en español del siguiente texto:\n\n${text}`;
        const apiKey = 'AIzaSyB_xgklkLXuE03VaibjWzT2kvNLxlwnims';         // <-- Reemplazar por tu clave real
        const model  = 'gemini-1.5-flash';

        const response = await fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{ parts: [{ text: prompt }] }]
            })
          }
        );

        if (!response.ok) throw new Error(await response.text());

        const data = await response.json();
        const corrected = data.candidates?.[0]?.content?.parts?.[0]?.text || 'Sin respuesta';

        selection.insertText(corrected, Word.InsertLocation.replace);
        document.getElementById('results').textContent = 'Corrección aplicada.';
        await context.sync();
      });
    } catch (error) {
      console.error(error);
      document.getElementById('results').textContent = 'Error: ' + error.message;
    }
  }
})();