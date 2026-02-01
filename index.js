const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.get("/", (req, res) => {
  res.send("Excel fill service running");
});

app.post(
  "/fill",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "payload", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      if (!req.files?.template || !req.files?.payload) {
        return res.status(400).send("Missing template or payload");
      }

      const templateBuffer = req.files.template[0].buffer;
      const payloadBuffer = req.files.payload[0].buffer;
      const payload = JSON.parse(payloadBuffer.toString("utf8"));

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(templateBuffer);

      const sheet = workbook.worksheets[0];

      // Beispiel: einfache Felder
      if (payload.customer_name) sheet.getCell("B2").value = payload.customer_name;
      if (payload.address) sheet.getCell("B3").value = payload.address;
      if (payload.project) sheet.getCell("B4").value = payload.project;

      const outBuffer = await workbook.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=filled.xlsx"
      );

      res.send(Buffer.from(outBuffer));
    } catch (err) {
      console.error(err);
      res.status(500).send("Internal error");
    }
  }
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Excel fill service running on port", PORT);
});
