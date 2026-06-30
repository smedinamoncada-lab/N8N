// ============================================================
// Microservicio gratuito para rellenar PDFs y planillas Excel
// (sustituye a PDF.co) - Desplegado gratis en Render.com
// ============================================================
//
// Endpoints disponibles:
//   POST /fill        -> rellena un PDF con campos de formulario
//   POST /fill-xlsx    -> rellena una planilla Excel celda por celda
//   POST /inspect      -> diagnóstico: lista los campos reales de un PDF
//   GET  /             -> healthcheck ("PDF filler activo")
// ============================================================

const express = require('express');
const { PDFDocument, PDFName } = require('pdf-lib');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.json({ limit: '25mb' }));

function isTruthy(v) {
  return v === true || v === 'True' || v === 'true' || v === '1';
}

// Algunos formularios (como BBVA y Banco Caja Social) implementan un
// "grupo de opciones" como UN SOLO campo checkbox que internamente tiene
// varias casillas (widgets/"kids"), una por opción - en vez de un
// verdadero grupo de radio buttons. pdf-lib no tiene un método de alto
// nivel para esto, así que seleccionamos manualmente la casilla correcta
// marcando su estado de apariencia (AS) y apagando las demás.
function selectCheckboxOption(form, groupName, optionName) {
  const field = form.getField(groupName);
  const widgets = field.acroField.getWidgets();
  const target = PDFName.of(optionName);
  let matched = false;

  for (const widget of widgets) {
    const apDict = widget.dict.lookup(PDFName.of('AP'));
    const nDict = apDict && apDict.lookup(PDFName.of('N'));
    const hasOption = !!(nDict && typeof nDict.keys === 'function' && nDict.keys().includes(target));
    if (hasOption) matched = true;
    widget.dict.set(PDFName.of('AS'), hasOption ? target : PDFName.of('Off'));
  }

  if (matched) {
    field.acroField.dict.set(PDFName.of('V'), target);
  }

  return matched;
}

// ============================================================
// POST /fill - Rellenar un PDF
// Body (JSON):
//   {
//     "pdfBase64": "....",
//     "fields": [
//       { "name": "Texto1", "value": "Juan Pérez" },
//       { "name": "estado civil.1", "value": "True" },   // BBVA/BCS: grupo.opcion
//       { "name": "TipoIdentificacion", "value": "0" },   // Banco Bogotá: radio group estándar
//       ...
//     ]
//   }
// ============================================================
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

      // 2) ¿Es un checkbox simple (un solo botón, sin opciones)?
      try {
        const cb = form.getCheckBox(name);
        if (isTruthy(value)) cb.check(); else cb.uncheck();
        continue;
      } catch (e) { /* no es checkbox simple */ }

      // 2b) ¿Es un grupo de radio buttons ESTÁNDAR? (nombre sin punto,
      //     value ya es la opción a seleccionar, ej "0","1","2"...)
      try {
        const rg = form.getRadioGroup(name);
        rg.select(String(value));
        continue;
      } catch (e) { /* no es radio group estándar */ }

      // 3) ¿Es una opción dentro de un grupo? ("grupo.opcion")
      const dotIndex = name.lastIndexOf('.');
      if (dotIndex > -1) {
        const groupName = name.substring(0, dotIndex);
        const optionName = name.substring(dotIndex + 1);

        // 3a) ¿Es un verdadero grupo de radio buttons?
        try {
          const rg = form.getRadioGroup(groupName);
          if (isTruthy(value)) rg.select(optionName);
          continue;
        } catch (e) { /* tampoco es radio group */ }

        // 3b) ¿Es un checkbox con varias casillas/opciones bajo el mismo nombre?
        try {
          form.getCheckBox(groupName); // valida que el campo exista y sea checkbox
          if (isTruthy(value)) {
            const ok = selectCheckboxOption(form, groupName, optionName);
            if (!ok) {
              warnings.push(`${name} (la opción "${optionName}" no existe en "${groupName}")`);
            }
          }
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

// ============================================================
// POST /fill-xlsx - Rellenar una planilla Excel celda por celda
// Body (JSON):
//   {
//     "xlsxBase64": "....",
//     "sheetName": "PLANILLA RADICACION",
//     "cells": [
//       { "cell": "E7", "value": "150,000,000" },
//       { "cell": "E88", "value": "Empresa de Prueba SAS" },
//       ...
//     ]
//   }
// ============================================================
app.post('/fill-xlsx', async (req, res) => {
  try {
    const { xlsxBase64, sheetName, cells } = req.body;

    if (!xlsxBase64 || !Array.isArray(cells)) {
      return res.status(400).json({
        error: 'Se requiere "xlsxBase64" (string) y "cells" (array de {cell, value}).',
      });
    }

    const buffer = Buffer.from(xlsxBase64, 'base64');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const worksheet = sheetName
      ? workbook.getWorksheet(sheetName)
      : workbook.worksheets[0];

    if (!worksheet) {
      return res.status(400).json({ error: `No se encontró la hoja "${sheetName}".` });
    }

    const warnings = [];

    for (const { cell, value } of cells) {
      if (value === null || value === undefined || value === '') continue; // no pisar celdas sin valor
      try {
        worksheet.getCell(cell).value = value;
      } catch (e) {
        warnings.push(`${cell}: ${e.message}`);
      }
    }

    const outBuffer = await workbook.xlsx.writeBuffer();

    res.json({
      xlsxBase64: Buffer.from(outBuffer).toString('base64'),
      warnings,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ============================================================
// POST /inspect - Diagnóstico: lista todos los campos de un PDF,
// y para los checkboxes con varias casillas, los nombres exactos
// de cada opción disponible.
// Body (JSON): { "pdfBase64": "...." }
// ============================================================
app.post('/inspect', async (req, res) => {
  try {
    const { pdfBase64 } = req.body;
    if (!pdfBase64) {
      return res.status(400).json({ error: 'Se requiere "pdfBase64" (string).' });
    }

    const pdfBytes = Buffer.from(pdfBase64, 'base64');
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const form = pdfDoc.getForm();

    const fields = form.getFields().map((f) => {
      const info = { name: f.getName(), type: f.constructor.name };

      if (typeof f.getOptions === 'function') {
        try {
          info.options = f.getOptions();
        } catch (e) { /* no aplica */ }
      }

      if (f.constructor.name === 'PDFCheckBox') {
        try {
          const widgets = f.acroField.getWidgets();
          info.widgetCount = widgets.length;
          info.optionsPerWidget = widgets.map((w) => {
            const apDict = w.dict.lookup(PDFName.of('AP'));
            const nDict = apDict && apDict.lookup(PDFName.of('N'));
            if (nDict && typeof nDict.keys === 'function') {
              return nDict
                .keys()
                .map((k) => k.toString().replace(/^\//, ''))
                .filter((k) => k !== 'Off');
            }
            return [];
          });
        } catch (e) {
          info.widgetError = e.message;
        }
      }

      return info;
    });

    res.json({ totalFields: fields.length, fields });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.send('PDF filler activo ✅'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor escuchando en puerto ${PORT}`));
