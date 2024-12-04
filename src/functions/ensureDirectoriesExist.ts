import fs from "fs";
import path from "path";

export async function ensureDirectoriesExist(basePath: string) {
  //   const directories = ["valid", "reports", "bad"];
  try {
    await fs.promises.mkdir(path.join(basePath, "data"), { recursive: true });
    console.log(`Directory data is ready.`);
  } catch (error) {
    console.error(`Error creating directory data:`, error);
  }

  //   for (const dir of directories) {
  //     const dirPath = path.join(basePath, dir);
  //     try {
  //       await fs.promises.mkdir(dirPath, { recursive: true });
  //       console.log(`Directory ${dirPath} is ready.`);
  //     } catch (error) {
  //       console.error(`Error creating directory ${dirPath}:`, error);
  //     }
  //   }
}
