import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

export async function convertCsvToExcel(
  csvFilePath: string,
  excelFilePath: string
) {
    await fs.promises.mkdir(path.join(__dirname, "..", "..", "csvs"), {
        recursive: true,
      });
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  // Read the CSV file
  const csvData = fs.readFileSync(csvFilePath, "utf8");

  // Split the CSV into rows
  const rows = csvData.split("\n").map((row) => row.split(","));

  // Add rows to the worksheet
  rows.forEach((row) => {
    worksheet.addRow(row);
  });

  // Write to an Excel file
  await workbook.xlsx.writeFile(`./csvs/${excelFilePath}-${Date.now()}.xlsx`);
  console.log(`CSV successfully converted to Excel at ${excelFilePath}`);
}
