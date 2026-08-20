import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "C:/Users/60893/Downloads/Production_KBMS_issue log_202511203 (2).xlsx";
const outputDir = "D:/CIE/BMS/outputs/019ffe8a-a18e-7cd2-bea0-9c45711a60e1";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));

const issueSheet = workbook.worksheets.getItem("ISSUE");
const rows = issueSheet.getRange("A1:K72").values;
const target = rows.slice(1).filter(r => ["ISSUE", "CR", "ADJUST"].includes(String(r[1] ?? "").trim().toUpperCase())).map(r => ({
  id: r[0], type: r[1], page: r[2], detail: String(r[3] ?? "").slice(0, 220),
  priority: r[4], status: r[7], remark: String(r[9] ?? r[10] ?? "").slice(0, 320),
}));
console.log("TARGET_ROWS", JSON.stringify(target));

for (let i = 0; i < workbook.worksheets.items.length; i++) {
  const sheet = workbook.worksheets.getItemAt(i);
  const used = sheet.getUsedRange();
  console.log(`USED ${sheet.name}: ${used?.address ?? "none"}`);
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 0.8, format: "png" });
  await fs.writeFile(`${outputDir}/source_${i + 1}.png`, new Uint8Array(await preview.arrayBuffer()));
}
