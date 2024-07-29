const fs = require("fs");
const pdfParse = require("pdf-parse");
const ExcelJS = require("exceljs");
const path = require("path");
const exec = require("child_process").exec;

const AbsPath = path.resolve();
const FolderPath = AbsPath + "\\archivos\\";
const file = "tarifa1"
main();

async function main() {
  const pdfData = await extractTextFromPDF(FolderPath + file + ".pdf");
  const lines = pdfData.text.split("\n");
  for (const line of lines) {
    await writeDataToExcel(line);
  }

  await workbookResumen.xlsx.writeFile(FolderPath + file + ".xlsx");
  exec(`start "" "${FolderPath + file + ".xlsx"}"`);
}

async function extractTextFromPDF(pdfPath) {
  const dataBuffer = fs.readFileSync(pdfPath);
  return await pdfParse(dataBuffer, { pagerender: render_page, version: "v2.0.550" });
}

function render_page(pageData) {
  let render_options = {
    normalizeWhitespace: true,
    disableCombineTextItems: true,
  };
  return pageData.getTextContent(render_options).then(function (textContent) {
    let lastY,
      text = "";
    for (let item of textContent.items) {
      if (lastY == item.transform[5] || !lastY) {
        text += "xYYx" + item.str;
      } else {
        text += "\n" + item.str;
      }
      lastY = item.transform[5];
    }
    return text;
  });
}

const workbookResumen = new ExcelJS.Workbook();
const worksheetResumen = workbookResumen.addWorksheet("Resumen");

async function writeDataToExcel(data) {

  worksheetResumen.addRow(data.split("xYYx"));
}
