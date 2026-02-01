const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");

const app = express();

// Multer: akzeptiert Uploads im Speicher (keine Disk)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB
});

// Healthcheck (damit du sehen kannst, ob Service läuft)
app.get("/", (req, res) => {
  res.status(200).send("OK - excel-fill-service is running");
});

/**
 * Erwartet multipart/form-data:
 * - file: XLSX Template (Binary)
 * - payload: JSON-String
 *
 * payload Beispiel:
 * {
 *   "sheetName": "Tabelle1",
 *   "cells": { "F5": "A26101", "B5": "Max Mustermann", "B6": "Musterstraße 1", "B7": "Projekt XY" },
 *   "positionStartRow": 15,
 *   "positions": [
 *     { "pos": 1, "title": "Spiegel", "desc": "Antikspiegel", "qty": "2 Stk", "dim": "69x90" }
 *   ],
 *   "columns": { "pos": "A", "title": "B", "desc": "C", "qty": "D", "dim": "E" }
 * }
 */
app.post("/fill", upload.fields([{ name: "file", maxCount: 1 }, { name: "payload", maxCount: 1 }]), async (req, res) => {
  try {
    const fileObj = req.files?.file?.[0];
if (!fileObj?.buffer) {
  return res.status(400).json({ error: "Missing file field 'file' (xlsx)." });
}
    }
    if (!req.body || !req.body.payload) {
      return res.status(400).json({ error: "Missing field 'payload' (json string)." });
    }

    let payload;
    try {
      payload = JSON.parse(req.body.payload);
    } catch {
      return res.status(400).json({ error: "Field 'payload' is not valid JSON." });
    }

    const sheetName = payload.sheetName || "Tabelle1";
    const cells = payload.cells || {};
    const positionStartRow = Number(payload.positionStartRow || 15);
    const positions = Array.isArray(payload.positions) ? payload.positions : [];

    // Spalten-Zuordnung für Positionszeile (Standard: A-E)
    const columns = payload.columns || { pos: "A", title: "B", desc: "C", qty: "D", dim: "E" };

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const ws = workbook.getWorksheet(sheetName) || workbook.worksheets[0];
    if (!ws) return res.status(400).json({ error: "Worksheet not found." });

    // 1) Fixe Zellen (Platzhalter) setzen
    for (const [address, value] of Object.entries(cells)) {
      ws.getCell(address).value = value ?? "";
    }

    // 2) Positionen ab positionStartRow schreiben
    // Hinweis: Wir überschreiben nur die Zellen in diesen Spalten – Formeln/Layout in anderen Spalten bleibt unangetastet.
    for (let i = 0; i < positions.length; i++) {
      const r = positionStartRow + i;
      const p = positions[i] || {};

      ws.getCell(`${columns.pos}${r}`).value = p.pos ?? (i + 1);
      ws.getCell(`${columns.title}${r}`).value = p.title ?? "";
      ws.getCell(`${columns.desc}${r}`).value = p.desc ?? "";
      ws.getCell(`${columns.qty}${r}`).value = p.qty ?? "";
      ws.getCell(`${columns.dim}${r}`).value = p.dim ?? "";
    }

    // XLSX zurückgeben
    const out = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="angebot.xlsx"');
    return res.status(200).send(Buffer.from(out));
  } catch (err) {
    return res.status(500).json({ error: "Server error", details: String(err?.message || err) });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`excel-fill-service listening on ${PORT}`));
