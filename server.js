// ============================================================
// Microservicio gratuito para rellenar el PDF (sustituye a PDF.co)
// Pensado para desplegarse gratis en Render.com (u otro hosting
// gratuito de Node.js: Railway, Fly.io, Cyclic, Vercel, etc.)
// ============================================================
//
// Endpoint: POST /fill
// Body (JSON):
//   {
//     "pdfBase64": "....",      // el PDF plantilla en base64
//     "fields": [               // mismos campos que ya usabas en PDF.co
//       { "name": "Día", "value": "21" },
//       { "name": "Vinculación inicial.1", "value": "True" },
//       ...
//     ]
//   }
//
// Respuesta:
//   { "pdfBase64": "...PDF relleno en base64...", "warnings": ["campo_x"] }
// ============================================================

const express = require('express');
const { PDFDocument } = require('pdf-lib');

const app = express();
app.use(express.json({ limit: '25mb' }));

function isTruthy(v) {
  return v === true || v === 'True' || v === 'true' || v === '1';
}

app.post('/fill', async (req, res) => {
  try {
    const { pdfBase64, fields } = req.body;

    if (!pdfBase64 || !Array.isArray(fields)) {
      return res.status(400).json({
        error: 'Se requiere "pdfBase64" (string) y "fields" (array de {name, value}).',
      });
    }

    const pdfBytes = Buffer.from(pdfBase64, 'base64');
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const form = pdfDoc.getForm();

    const warnings = [];

    for (const { name, value } of fields) {
      // 1) ¿Es un campo de texto?
      try {
        const tf = form.getTextField(name);
        tf.setText(value === null || value === undefined ? '' : String(value));
        continue;
      } catch (e) { /* no es text field */ }

      // 2) ¿Es un checkbox simple?
      try {
        const cb = form.getCheckBox(name);
        if (isTruthy(value)) cb.check(); else cb.uncheck();
        continue;
      } catch (e) { /* no es checkbox */ }

      // 3) ¿Es una opción de un grupo de radio buttons? ("grupo.opcion")
      const dotIndex = name.lastIndexOf('.');
      if (dotIndex > -1) {
        const groupName = name.substring(0, dotIndex);
        const optionName = name.substring(dotIndex + 1);

        try {
          const rg = form.getRadioGroup(groupName);
          if (isTruthy(value)) rg.select(optionName);
          continue;
        } catch (e) { /* tampoco es radio group */ }

        try {
          const cb2 = form.getCheckBox(groupName);
          if (isTruthy(value)) cb2.check(); else cb2.uncheck();
          continue;
        } catch (e) { /* nada coincide */ }
      }

      warnings.push(name);
    }

    const filledBytes = await pdfDoc.save();

    res.json({
      pdfBase64: Buffer.from(filledBytes).toString('base64'),
      warnings,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.send('PDF filler activo ✅'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor escuchando en puerto ${PORT}`));
