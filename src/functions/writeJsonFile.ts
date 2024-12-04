import fs from "fs";

export async function writeJsonFile(filePath: string, data: any) {
  try {
    const jsonData = JSON.stringify(data, null, 1);
    await fs.promises.writeFile(filePath, jsonData);
    console.log(`report successfully written to ${filePath}`);
  } catch (error) {
    console.error("reportJSON file:", error);
  }
}
