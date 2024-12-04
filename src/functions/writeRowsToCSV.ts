import ExcelJS from "exceljs";

export async function writeRowsToCSV(filePath: string, rows: ExcelJS.RowMap[]) {
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
