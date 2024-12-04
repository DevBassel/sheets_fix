import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

export async function convertExcelToCSV(
  inputFilePath: string,
  outputFilePath: string
) {
  await fs.promises.mkdir(path.join(__dirname, "..", "..", "csvs"), {
    recursive: true,
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFilePath);

  const worksheet = workbook.getWorksheet(1); // Get the first sheet

  if (!worksheet) {
    console.log("No worksheet found.");
    return;
  }

  const csvData: string[] = [];

  // Iterate through all rows in the worksheet
  worksheet.eachRow((row: ExcelJS.RowMap) => {
    const rowData = row.values
      .slice(1) // Skip the first index as it's undefined
      .map((cell: any) => {
        if (cell == "31612121605434") {
          console.log("TCL: row", row.values);
        }

        if (typeof cell == "object") {
          if (!!cell.richText) {
            return `"${cell.richText.map((t: any) => t.text).join(" ")}"`
              .replace(/[\r\n]+/g, "")
              .trim();
          } else {
            if (!cell.result) {
              return `""`.replace(/[\r\n]+/g, "").trim();
            }
          }
        } else {
          return `"${cell}"`.replace(/[\r\n]+/g, "").trim();
        }
      })
      .join(",");
    csvData.push(rowData);
  });

  // Write the CSV data to a file
  fs.writeFileSync(
    `./csvs/${outputFilePath}-${Date.now()}-.csv`,
    csvData.join("\n"),
    "utf8"
  );
  console.log(`File successfully converted to CSV at ${outputFilePath}`);
}
