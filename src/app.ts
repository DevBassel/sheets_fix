import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

interface IRep {
  badDataCount: number;
  goodDataCount: number;
  gov: any[];
  mange: any[];
  center: any[];
  classes: any[];
}

async function openExcelFile(
  filePath: string,
  outFileName: string,
  type: "CSV" | "xlsx"
) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    //   validate
    const outValidPath = `./data/valid/valide-${outFileName}-${Date.now().toString()}.${type}`;
    const validRows: ExcelJS.RowMap[] = [];
    // bad data
    const outBadPath = `./data/bad/bad-${outFileName}-${Date.now()}.${type}`;
    const badRows: ExcelJS.RowMap[] = [];

    // report
    const report: IRep = {
      gov: [], // row index = 3
      mange: [], // row index = 4
      center: [], // row index = 5
      classes: [],
      badDataCount: 0,
      goodDataCount: 0,
    };
    const reportPath = `./data/reports/rep-${outFileName}-${Date.now()}.json`;

    const worksheet = workbook.getWorksheet(1);
    if (worksheet) {
      const headerRow = worksheet.getRow(1);
      badRows.push(headerRow);

      worksheet.eachRow((row) => {
        let good = 0;

        row.eachCell((cell, colNumber) => {
          if (cell.value) {
            let cellVal = cell.value as any;
            if (cellVal.result) {
              cellVal = cellVal.result;
            }
            switch (colNumber) {
              // name col
              case 1:
                good++;
                break;

              // school col
              case 2:
                good++;
                break;

              // gov col
              case 3:
                if (!report.gov.includes(cellVal)) report.gov.push(cellVal);
                good++;
                break;

              // mange
              case 4:
                if (!report.mange.includes(cellVal)) report.mange.push(cellVal);
                good++;
                break;

              // center
              case 5:
                if (!report.center.includes(cellVal))
                  report.center.push(cellVal);
                good++;
                break;

              // vaillage col
              case 6:
                good++;
                break;

              // nid col
              case 7:
                good++;
                break;

              // gender col
              case 8:
                good++;
                break;

              // address col
              case 9:
                good++;
                break;

              // phone col
              case 10:
                if (cellVal.length > 9) good++;
                break;

              // class col
              case 11:
                if (
                  !report.classes.includes(cellVal) &&
                  typeof cellVal != "object"
                ) {
                  report.classes.push(cellVal);
                  good++;
                }
                break;
            }
          }
        });

        if (good == 11) validRows.push(row);
        else badRows.push(row);

        good = 0;
      });

      // result
      report.badDataCount = badRows.length;
      report.goodDataCount = validRows.length;

      if (type == "CSV") {
        await writeRowsToCSV(outValidPath, validRows);
        await writeRowsToCSV(outBadPath, badRows);
      } else {
        await writeRowsToNewExcel(outValidPath, validRows);
        await writeRowsToNewExcel(outBadPath, badRows);
      }
      await writeJsonFile(reportPath, report);
    }
  } catch (error) {
    console.log("TCL: error", error);
  }
}

// write new sheet
async function writeRowsToNewExcel(filePath: string, rows: ExcelJS.RowMap[]) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  rows.forEach((row) => {
    if (row.values) {
      const rowData = row.values.slice(1);
      const formattedRow = rowData.map((cell: any) => {
        return `${cell}`;
      });
      worksheet.addRow(formattedRow);
    }
  });

  await workbook.xlsx.writeFile(filePath);
  console.log(`Rows written successfully to ${filePath}`);
}

// write CSV
async function writeRowsToCSV(filePath: string, rows: ExcelJS.RowMap[]) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  // Add rows to the worksheet
  rows.forEach((row) => {
    if (row.values) {
      const rowData = row.values.slice(1); // Skipping the first index (if necessary)
      const formattedRow = rowData.map((cell: any) => {
        // Ensure that cell values are converted to string format
        return `${cell}`;
      });
      worksheet.addRow(formattedRow);
    }
  });

  // Export the data as CSV
  await workbook.csv.writeFile(filePath);
  console.log(`Rows written successfully to ${filePath}`);
}

async function writeJsonFile(filePath: string, data: any) {
  try {
    const jsonData = JSON.stringify(data, null, 1);
    await fs.promises.writeFile(filePath, jsonData);
    console.log(`report successfully written to ${filePath}`);
  } catch (error) {
    console.error("reportJSON file:", error);
  }
}

async function ensureDirectoriesExist(basePath: string) {
  const directories = ["valid", "reports", "bad"];

  for (const dir of directories) {
    const dirPath = path.join(basePath, dir);
    try {
      await fs.promises.mkdir(dirPath, { recursive: true });
      console.log(`Directory ${dirPath} is ready.`);
    } catch (error) {
      console.error(`Error creating directory ${dirPath}:`, error);
    }
  }
}

function checkValue(val: string) {}
const basePath = "./data";
ensureDirectoriesExist(basePath);

openExcelFile("./data/زفتي.xlsx", "زفتي", "xlsx");
