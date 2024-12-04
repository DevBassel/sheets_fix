import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { IRep } from "./types/report.interface";
import { ensureDirectoriesExist } from "./functions/ensureDirectoriesExist";
import { writeRowsToCSV } from "./functions/writeRowsToCSV";
import { writeRowsToNewExcel } from "./functions/writeRowsToNewExcel";
import { writeJsonFile } from "./functions/writeJsonFile";
import { convertExcelToCSV } from "./functions/convertExcelToCSV";
import { convertCsvToExcel } from "./functions/convertCsvToExcel";

async function openExcelFile(filePath: string, outFileName: string) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const resultPath = path.join(
      __dirname,
      "..",
      "data",
      `${outFileName}-${Date.now().toString()}`
    );
    await fs.promises.mkdir(path.join(resultPath), { recursive: true });

    // good data
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
      cls: {},
    };

    let cls1: string[] = [];
    let cls2: string[] = [];
    let cls3: string[] = [];
    let cls4: string[] = [];
    let cls5: string[] = [];
    let cls6: string[] = [];

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
                if (typeof cellVal == "object") {
                  console.log("TCL: openExcelFile -> cellVal", cellVal);
                }
                good++;
              }
              break;
            // school
            case 2:
              if (checkValue(cellVal)) {
                good++;
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
                if (/(ول|1)/gi.test(cellVal)) {
                  cls1.push(cellVal);
                }
                if (/(ثان|2)/gi.test(cellVal)) {
                  cls2.push(cellVal);
                }
                if (/(ثال|3)/gi.test(cellVal)) {
                  cls3.push(cellVal);
                }
                if (/(را|4)/gi.test(cellVal)) {
                  cls4.push(cellVal);
                }
                if (/(خا|5)/gi.test(cellVal)) {
                  cls5.push(cellVal);
                }
                if (/(سا|6)/gi.test(cellVal)) {
                  cls6.push(cellVal);
                }
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
      report.cls["class1"] = cls1;
      report.cls["class2"] = cls2;
      report.cls["class3"] = cls3;
      report.cls["class4"] = cls4;
      report.cls["class5"] = cls5;
      report.cls["class6"] = cls6;
      // result
      report.badDataCount = badRows.length;
      report.goodDataCount = validRows.length;

      await writeRowsToCSV(outValidPath.replace(".xlsx", ".csv"), validRows);
      await writeRowsToNewExcel(outValidPath, validRows);
      await writeRowsToNewExcel(outBadPath, badRows);

      await writeJsonFile(path.join(resultPath, `Report.json`), report);
    }
  } catch (error) {
    console.log("TCL: error", error);
  }
}

function checkValue(val: string) {
  return !!val;
}

ensureDirectoriesExist("./");

// open file and work
// openExcelFile("./clean_sheets/غرب طنطا.xlsx", "غرب طنطا");
// convertCsvToExcel("./csvs/غرب طنطا-1733289544445-.csv", "غرب طنطا");

// to remove all styles from fken sheet output ==> ./csvs
// convertExcelToCSV("./sheets/غرب طنطا.xlsx", "غرب طنطا");
