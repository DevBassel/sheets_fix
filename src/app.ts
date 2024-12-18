import { convertExcelToCSV } from "./functions/convertExcelToCSV";
import { ensureDirectoriesExist } from "./functions/ensureDirectoriesExist";
import { handelStuSheets } from "./functions/handelStuSheets";

ensureDirectoriesExist("./");

// open file and work
// handelStuSheets("./clean_sheets/رسمي لغات.xlsx", "رسمي لغات"); // done
// handelStuSheets("./clean_sheets/رسمي حكومي.xlsx", "رسمي حكومي"); // done
// handelStuSheets("./clean_sheets/زفتي.xlsx", "زفتي");
// handelStuSheets("./clean_sheets/شرق طنطا.xlsx", "شرق طنطا");
// handelStuSheets("./clean_sheets/غرب طنطا.xlsx", "غرب طنطا");
// handelStuSheets("./clean_sheets/new_1.xlsx", "new_1");
// handelStuSheets("./clean_sheets/new_2.xlsx", "new_2");

// convertCsvToExcel("./csvs/غرب طنطا-1733289544445-.csv", "غرب طنطا");

// to remove all styles from fken sheet output ==> ./csvs
// convertExcelToCSV("./clean_sheets/الممرضات.xlsx", "الممرضات");
