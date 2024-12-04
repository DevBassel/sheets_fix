import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { IRep } from "./types/report.interface";
import { ensureDirectoriesExist } from "./functions/ensureDirectoriesExist";
import { writeRowsToCSV } from "./functions/writeRowsToCSV";
import { writeRowsToNewExcel } from "./functions/writeRowsToNewExcel";
import { writeJsonFile } from "./functions/writeJsonFile";
import { convertExcelToCSV } from "./functions/convertExcelToCSV";

async function openExcelFile(filePath: string, outFileName: string) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    //   validate
    const resultPath = path.join(
      __dirname,
      "..",
      "data",
      `${outFileName}-${Date.now().toString()}`
    );

    await fs.promises.mkdir(path.join(resultPath), { recursive: true });

    const outValidPath = path.join(resultPath, `Good.xlsx`);
    const validRows: ExcelJS.RowMap[] = [];
    // bad data
    const outBadPath = path.join(resultPath, `Bad.xlsx`);
    const badRows: ExcelJS.RowMap[] = [];

    // report
    const report: IRep = {
      gov: [],
      mange: [],
      center: [],
      classes: [],
      badDataCount: 0,
      goodDataCount: 0,
    };
    const reportPath = path.join(resultPath, `Report.json`);

    const worksheet = workbook.getWorksheet(1);
    if (worksheet) {
      const headerRow = worksheet.getRow(1);
      badRows.push(headerRow);

      worksheet.eachRow((row) => {
        let good = 0;

        row.eachCell((cell, colNumber) => {
          cell.style = {};
          let cellVal = cell.value as any;
          if (cellVal.result) {
            cellVal = cellVal.result;
          }
          switch (colNumber) {
            // name
            case 1:
              if (checkValue(cellVal)) {
                good++;
              }
              break;
            // school
            case 2:
              if (checkValue(cellVal)) {
                good++;
              }
              if (typeof cellVal === "object") {
                console.log("TCL: cellVal", cellVal);
              }
              break;
            // gov col
            case 3:
              if (!report.gov.includes(cellVal)) report.gov.push(cellVal);
              if (checkValue(cellVal)) {
                good++;
              }
              break;

            // mange
            case 4:
              if (!report.mange.includes(cellVal)) report.mange.push(cellVal);
              if (checkValue(cellVal)) {
                good++;
              }
              break;

            // center
            case 5:
              if (!report.center.includes(cellVal)) report.center.push(cellVal);
              if (checkValue(cellVal)) {
                good++;
              }
              break;
            // nid
            case 7:
              if (checkValue(cellVal)) {
                good++;
              }
              break;
            // phone
            case 10:
              if (checkValue(cellVal)) {
                good++;
              }
              break;
            // class col
            case 11:
              if (
                !report.classes.includes(cellVal) &&
                typeof cellVal != "object"
              ) {
                report.classes.push(cellVal);
              }
              if (checkValue(cellVal)) {
                good++;
              }
              break;
          }
        });

        if (good == 8) validRows.push(row);
        else badRows.push(row);

        good = 0;
      });

      // result
      report.badDataCount = badRows.length;
      report.goodDataCount = validRows.length;

      await writeRowsToCSV(outValidPath.replace(".xlsx", ".csv"), validRows);
      await writeRowsToNewExcel(outValidPath, validRows);
      await writeRowsToNewExcel(outBadPath, badRows);

      await writeJsonFile(reportPath, report);
    }
  } catch (error) {
    console.log("TCL: error", error);
  }
}

function checkValue(val: string) {
  return !!val;
}

ensureDirectoriesExist("./");

openExcelFile("./clean_sheets/زفتي.xlsx", "زفتي");


// to remove all styles from fken sheet
// convertExcelToCSV("./sheets/زفت22ي.xlsx", "زفتي");
