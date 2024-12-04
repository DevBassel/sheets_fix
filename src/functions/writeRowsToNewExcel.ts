import ExcelJS from "exceljs";

export async function writeRowsToNewExcel(
  filePath: string,
  rows: ExcelJS.RowMap[]
) {
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
