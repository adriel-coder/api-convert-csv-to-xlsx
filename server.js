const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");

const app = express();

// aumenta limite pra CSVs maiores
app.use(express.json({ limit: "20mb" }));
app.use(cors());

/**
 * POST /convert
 * body:
 * {
 *   "csvText": "id,name\n1,Ana\n2,Bob"
 *   // OU
 *   "csvBase64": "aWQsbmFtZQoxLEFuYQoyLEJvYg=="
 *   "sheetName": "Planilha" (opcional)
 *   "fileName": "planilha.xlsx" (opcional)
 *   "return": "base64" | "dataUri" | "file" (opcional; default base64)
 * }
 */
app.post("/convert", async (req, res) => {
  let csvText = String(req.body.csvText ?? "");

  if (!csvText.includes("\n")) {
    // tenta recuperar quando veio tudo em uma linha com espaços
    // (assumindo que cada linha começa com número ou com o header)
    csvText = csvText
      .replace(/\s+(?=\d+,)/g, "\n") // antes de "1,", "2,", ...
      .replace(/\s+(?=id,)/i, "\n"); // se por algum motivo header não estiver no começo
    csvText = csvText.trim();
  }
  console.log(csvText);

  // 1) Converte CSV texto -> worksheet
  const wb = XLSX.read(csvText, { type: "string" });
  const oldName = wb.SheetNames[0];
  const newName = "Planilha";

  const ws0 = wb.Sheets[wb.SheetNames[0]];
  console.log("RANGE:", ws0?.["!ref"]); // esperado: A1:E4

  wb.Sheets[newName] = wb.Sheets[oldName];
  delete wb.Sheets[oldName];
  wb.SheetNames[0] = newName;

  // 3) Gera XLSX como bytes (buffer/uint8array)
  const xlsxBuffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  // 4) Converte bytes -> base64
  const xlsxBase64 = Buffer.from(xlsxBuffer).toString("base64");

  // 5) Data URI (muito aceito em campos FILE)
  const fileDataUri = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${xlsxBase64}`;

  res.json({ fileDataUri });

  //   try {
  //     const {
  //       csvText,
  //       csvBase64,
  //       sheetName = "Planilha",
  //       fileName = "planilha.xlsx",
  //       return: returnType = "base64",
  //     } = req.body ?? {};

  //     if (!csvText && !csvBase64) {
  //       return res.status(400).json({
  //         error: 'Envie "csvText" (string) ou "csvBase64" (string).',
  //       });
  //     }

  //     // 1) Obter o CSV como texto UTF-8
  //     const csv = csvText
  //       ? String(csvText)
  //       : Buffer.from(String(csvBase64), "base64").toString("utf-8");

  //     // 2) CSV -> Worksheet
  //     const ws = XLSX.read(csv, { type: "string" });

  //     // 3) Worksheet -> Workbook
  //     const wb = XLSX.utils.book_new();
  //     XLSX.utils.book_append_sheet(wb, ws, sheetName);

  //     // 4) Workbook -> XLSX bytes (Buffer)
  //     const xlsxBuffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  //     // 5) Resposta no formato desejado
  //     if (returnType === "file") {
  //       res.setHeader(
  //         "Content-Type",
  //         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  //       );
  //       res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
  //       return res.send(xlsxBuffer);
  //     }

  //     const base64 = Buffer.from(xlsxBuffer).toString("base64");
  //     if (returnType === "dataUri") {
  //       const dataUri =
  //         "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," +
  //         base64;
  //       return res.json({ fileName, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", dataUri });
  //     }

  //     // default: base64 puro
  //     return res.json({ fileName, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", base64 });
  //   } catch (err) {
  //     return res.status(500).json({
  //       error: "Falha ao converter CSV para XLSX",
  //       details: err?.message ?? String(err),
  //     });
  //   }
});

app.get("/", (req, res) => res.json({ ok: true }));

app.use((err, req, res, next) => {
  console.error("ERRO:", err);
  res.status(500).json({
    error: "Internal Server Error",
    details: String(err?.message ?? err),
  });
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`CSV->XLSX API rodando em http://localhost:${port}`);
});
