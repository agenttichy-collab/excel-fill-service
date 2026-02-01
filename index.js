import express from "express";
import cors from "cors";
import multer from "multer";
import ExcelJS from "exceljs";

const app = express();
app.use(cors());

// Multer: multipart/form-data robust in memory
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024, // 25MB
  },
});

// Health
app.get("/", (req, res) => {
  res.status(200).send("OK excel fill service running");
});

// Wichtig:
// - akzeptiert Datei unter "template" ODER "file"
// - akzeptiert payload als Textfeld ODER als Datei "payload"
const multipartHandler = upload.fields([
  { name: "template", maxCount: 1 },
  { name: "file", maxCount: 1 },
  { name: "payload", maxCount: 1 },
]);

app.post("/fill", (req, res) => {
  multipartHandler(req, res, async (err) => {
    try {
      if (err) {
        console.error("MULTER_ERROR:", err);
        return res.status(400).json({
          ok: false,
          error: "multipart_parse_error",
          detail: String(err?.message || err),
        });
      }

      // Datei finden (template oder file)
      const templateFile =
        (req.files?.template && req.files.template[0]) ||
        (req.files?.file && req.files.file[0]);

      if (!templateFile?.buffer) {
        console.error("NO_TEMPLATE_FILE. files:", Object.keys(req.files || {}));
        return res.status(400).json({
          ok: false,
          error: "missing_template_file",
          hint: 'Sende ein Binary-Feld "template" oder "file" (multipart/form-data).',
          gotFiles: Object.keys(req.files || {}),
        });
      }

      // Payload lesen:
      // 1) bevorzugt Textfeld req.body.payload
      // 2) alternativ payload als Datei (payload.json)
      let payloadRaw = req.body?.payload;

      if (!payloadRaw && req.files?.payload?.[0]?.buffer) {
        payloadRaw = req.files.payload[0].buffer.toString("utf8");
      }

      if (!payloadRaw) {
        console.error("NO_PAYLOAD. body keys:", Object.keys(req.body || {}));
        return res.status(400).json({
          ok: false,
          error: "missing_payload",
          hint: 'Sende ein Textfeld "payload" (JSON-String) ODER eine Datei "payload" (payload.json).',
          gotBodyKeys: Object.keys(req.body || {}),
          gotFiles: Object.keys(req.files || {}),
        });
      }

      let payload;
      try {
        payload = JSON.parse(payloadRaw);
      } catch (e) {
        console.error("PAYLOAD_JSON_PARSE_FAIL:", e);
        return res.status(400).json({
          ok: false,
          error: "payload_not_json",
          detail: String(e?.message || e),
          payloadPreview: payloadRaw.slice(0, 300),
        });
      }

      const sheetName = payload.sheetName || "Tabelle1";
      const cells = payload.cells || {}; // { "B5":"...", "F5":"..." }

      // Workbook laden
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(templateFile.buffer);

      const worksheet = workbook.getWorksheet(sheetName) || workbook.worksheets[0];
      if (!worksheet) {
        return res.status(400).json({
          ok: false,
          error: "worksheet_not_found",
          detail: `Blatt "${sheetName}" nicht gefunden`,
          availableSheets: workbook.worksheets.map((w) => w.name),
        });
      }

      // Zellen setzen
      for (const [addr, val] of Object.entries(cells)) {
        worksheet.getCell(addr).value = val ?? "";
      }

      // Optional: Positionsliste (wenn du sowas sendest)
      // payload.positions: [{ row: 15, values: { A:"1", B:"...", C:"..." } }, ...]
      if (Array.isArray(payload.positions)) {
        for (const pos of payload.positions) {
          if (!pos?.row || !pos?.values) continue;
          for (const [col, v] of Object.entries(pos.values)) {
            worksheet.getCell(`${col}${pos.row}`).value = v ?? "";
          }
        }
      }

      // zurück als Excel
      const outBuffer = await workbook.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader("Content-Disposition", 'attachment; filename="filled.xlsx"');

      return res.status(200).send(Buffer.from(outBuffer));
    } catch (e) {
      console.error("FILL_ERROR:", e);
      return res.status(500).json({
        ok: false,
        error: "internal_error",
        detail: String(e?.message || e),
      });
    }
  });
});

// Railway/Render nutzen PORT
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Server listening on", PORT));
