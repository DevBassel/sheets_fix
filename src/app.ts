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

    let cls1: Set<string> = new Set();
    let cls2: Set<string> = new Set();
    let cls3: Set<string> = new Set();
    let cls4: Set<string> = new Set();
    let cls5: Set<string> = new Set();
    let cls6: Set<string> = new Set();

    const worksheet = workbook.getWorksheet(1);
    if (worksheet) {
      const headerRow = worksheet.getRow(1);
      validRows.push(headerRow);

      worksheet.eachRow((row) => {
        let good = 0;

        row.eachCell((cell, colNumber) => {
          cell.style = {};
          let cellVal = cell.value as any;
          // let cellVal = handelCell(cell.value);
          if (cellVal.result) {
            cellVal = cellVal.result;
          }
          if (cellVal == "احمد هيثم محمد سعيد") {
            console.log("TCL: row", row.values);
          }

          switch (colNumber) {
            // name
            case 1:
              if (checkValue(cellVal)) {
                cell.value = `${String(cellVal)}`
                  .replace(/[\r\n]+/g, "")
                  .trim();
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
                if (/(غرب|زفت)/gi.test(cellVal)) {
                  cell.value = "محافظه الغربية";
                  good++;
                }
              }
              break;

            // mange
            case 4:
              if (!report.mange.includes(cellVal)) report.mange.push(cellVal);
              if (checkValue(cellVal)) {
                if (/(سنط)/gi.test(cellVal)) {
                  cell.value = "اداره السنطة";
                  good++;
                }

                if (/(زفت)/gi.test(cellVal)) {
                  cell.value = "ادارة زفتى";
                  good++;
                }

                if (/(شرق طن)/gi.test(cellVal)) {
                  cell.value = "ادارة شرق طنطا";
                  good++;
                }
                if (/(غرب طن)/gi.test(cellVal)) {
                  cell.value = "ادارة غرب طنطا";
                  good++;
                }
              }
              break;

            // center
            case 5:
              if (!report.center.includes(cellVal)) report.center.push(cellVal);
              if (checkValue(cellVal)) {
                if (/(سنط)/gi.test(cellVal)) {
                  cell.value = "مركز السنطة";
                  good++;
                }

                if (/(زفت)/gi.test(cellVal)) {
                  cell.value = "مركز زفتى";
                  good++;
                }

                if (/(طنطا|طنطا - غربية)/gi.test(cellVal)) {
                  cell.value = "مركز طنطا";
                  good++;
                }
              }
              break;
            // nid
            case 7:
              if (checkValue(cellVal)) {
                good++;
              } else {
                console.log("TCL: cellVal", cell.value);
              }
              break;
            // phone
            case 10:
              if (checkValue(cellVal)) {
                if (String(cellVal).length > 4) good++;
              }
              break;
            // class col
            case 11:
              if (!report.classes.includes(cellVal)) {
                report.classes.push(cellVal);
              }
              if (checkValue(cellVal)) {
                if (/(ول|1)/gi.test(cellVal)) {
                  cls1.add(cellVal);
                  cell.value = "الصف الاول الابتدائي";
                }
                if (/(ثان|2)/gi.test(cellVal)) {
                  cls2.add(cellVal);
                  cell.value = "الصف الثاني الابتدائي";
                }
                if (/(ثال|3)/gi.test(cellVal) || /(القالث)/gi.test(cellVal)) {
                  cls3.add(cellVal);
                  cell.value = "الصف الثالث الابتدائي";
                }
                if (/(راب|4)/gi.test(cellVal)) {
                  cls4.add(cellVal);
                  cell.value = "الصف الرابع الابتدائي";
                }
                if (/(امس|5)/gi.test(cellVal)) {
                  cls5.add(cellVal);
                  cell.value = "الصف الخامس الابتدائي";
                }
                if (/(سا|6)/gi.test(cellVal)) {
                  cls6.add(cellVal);
                  cell.value = "الصف السادس الابتدائي";
                }
                good++;
              }
              break;
          }
        });

        if (good == 8) validRows.push(row);
        else badRows.push(row);

        good = 0;
      });
      report.cls["class1"] = [...cls1];
      report.cls["class2"] = [...cls2];
      report.cls["class3"] = [...cls3];
      report.cls["class4"] = [...cls4];
      report.cls["class5"] = [...cls5];
      report.cls["class6"] = [...cls6];
      report.classes = [];
      // result
      report.badDataCount = badRows.length;
      report.goodDataCount = validRows.length;

      await writeRowsToCSV(outValidPath.replace(".xlsx", ".csv"), validRows);
      await writeRowsToCSV(outBadPath.replace(".xlsx", ".csv"), badRows);
      await writeRowsToNewExcel(outValidPath, validRows);
      await writeRowsToNewExcel(outBadPath, badRows);

      await writeJsonFile(path.join(resultPath, `Report.json`), report);
    }
  } catch (error) {}
}

function checkValue(val: string) {
  return !!val;
}

ensureDirectoriesExist("./");

// open file and work
// openExcelFile("./clean_sheets/رسمي لغات.xlsx", "رسمي لغات"); // done
// openExcelFile("./clean_sheets/رسمي حكومي.xlsx", "رسمي حكومي"); // done
// openExcelFile("./clean_sheets/زفتي.xlsx", "زفتي");
// openExcelFile("./clean_sheets/شرق طنطا.xlsx", "شرق طنطا");
openExcelFile("./clean_sheets/غرب طنطا.xlsx", "غرب طنطا");

// convertCsvToExcel("./csvs/غرب طنطا-1733289544445-.csv", "غرب طنطا");

// to remove all styles from fken sheet output ==> ./csvs
// convertExcelToCSV("./sheets/زفتي.xlsx", "زفتي");
