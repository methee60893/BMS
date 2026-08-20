import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath = "C:/Users/60893/Downloads/Production_KBMS_issue log_202511203 (2).xlsx";
const outputDir = "D:/CIE/BMS/outputs/019ffe8a-a18e-7cd2-bea0-9c45711a60e1";
const outputPath = `${outputDir}/KBMS_Sprint_Plan_Analysis.xlsx`;

const sourceWb = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const sourceRows = sourceWb.worksheets.getItem("ISSUE").getRange("A1:K72").values.slice(1);
const targetRows = sourceRows.filter(r => ["ISSUE", "CR", "ADJUST"].includes(String(r[1] ?? "").trim().toUpperCase()));

const text = v => String(v ?? "").trim();
const upper = v => text(v).toUpperCase();
const excelDate = v => typeof v === "number" && Number.isFinite(v) ? new Date(Date.UTC(1899, 11, 30) + v * 86400000) : v;
const signal = remark => {
  const r = text(remark).toLowerCase();
  if (r.includes("รีบ") || r.includes("high!!!!")) return "เร่งด่วนตาม Remark";
  if (r.includes("ดำเนินการ")) return "กำลังดำเนินการ";
  if (r.includes("confirm") || r.includes("ไม่แน่ใจ") || r.includes("พิจารณา")) return "ต้องยืนยัน/ตัดสินใจ";
  if (r.includes("investigate") || r.includes("ยาก")) return "ต้องวิเคราะห์ก่อนทำ";
  if (r.includes("monitor") || r.includes("follow-up")) return "ต้องติดตามผล";
  return "";
};

const sprintMap = new Map([
  [65, ["Sprint 1", 1, 5, "Blocker และเร่งด่วน; งานรองรับข้อมูลมากกว่า 300 บรรทัด ประเมิน Large รวม Verify"]],
  [68, ["Sprint 1", 2, 3, "Blocker และเร่งด่วน; แก้ Approve OTB พร้อม Regression Test"]],
  [37, ["Sprint 2", 1, 3, "Very High; เปลี่ยนการคำนวณ Export TXN และต้องตรวจผลรวม"]],
  [58, ["Sprint 2", 2, 3, "Very High; Requirement หลักของชุด Warning ใน Create Switch"]],
  [64, ["Sprint 2", 3, 2, "Very High; ทำต่อจาก #58 และใช้ชุดทดสอบร่วมกัน"]],
  [69, ["Sprint 3", 1, 3, "Very High; เพิ่ม Preview/Summary ก่อนส่ง SAP และ Verify ยอด"]],
  [12, ["Sprint 3", 2, 2, "High; เพิ่ม Total Actual + Draft PO และตรวจยอดรวม"]],
  [31, ["Sprint 3", 3, 3, "High; ต่อเนื่องจาก #12 ใน Export Summary พร้อม Regression"]],
  [9,  ["Sprint 4", 1, 3, "High; Feature หลัก Multi-filter บน Draft PO"]],
  [15, ["Sprint 4", 2, 2, "High; Reuse แนวทาง Multi-filter บน OTB Remaining"]],
  [16, ["Sprint 4", 3, 2, "High; Reuse แนวทางบน Draft OTB และทวน UAT เดิม"]],
  [17, ["Sprint 4", 4, 1, "High; Reuse แนวทางบน Approved OTB และทวน UAT เดิม"]],
  [38, ["Sprint 5", 1, 2, "High; เพิ่มข้อมูล Cate/Brand ใน Preview และตรวจกรณี Budget ไม่พอ"]],
  [52, ["Sprint 5", 2, 3, "Medium และกำลังดำเนินการ; เติมข้อมูลผลลัพธ์ Bulk Upload ให้ครบ"]],
  [8,  ["Sprint 5", 3, 3, "Medium; เติม Segment อัตโนมัติและตรวจ Mapping"]],
  [49, ["Sprint 6", 1, 5, "Medium แต่กระทบ Report หลายหน้า; ต้อง Regression Test กว้าง"]],
  [5,  ["Sprint 6", 2, 2, "Low; ปรับการแสดงผลและ Freeze Header"]],
  [6,  ["Sprint 7", 1, 3, "Low; ปรับสีหลายหน้าและตรวจความสม่ำเสมอ"]],
  [3,  ["Sprint 7", 2, 3, "Low แต่ไม่แน่นอน; รวม Technical Spike และ Verify ไว้ใน Estimate"]],
]);

const wb = Workbook.create();
const summary = wb.worksheets.add("Sprint Summary");
const plan = wb.worksheets.add("Sprint Plan");
const tbc = wb.worksheets.add("TBC & Decisions");
const all = wb.worksheets.add("All Analysis");
const rules = wb.worksheets.add("Planning Rules");
for (const s of wb.worksheets.items) s.showGridLines = false;

const navy = "#17365D", blue = "#1F4E78", white = "#FFFFFF", lightBlue = "#D9EAF7";
const orange = "#F4B183", paleOrange = "#FCE4D6", paleYellow = "#FFF2CC", paleGreen = "#E2F0D9";
const paleRed = "#F4CCCC", grey = "#E7E6E6", dark = "#1F1F1F";
const titleStyle = { fill: navy, font: { bold: true, color: white, size: 18 }, verticalAlignment: "center" };
const sectionStyle = { fill: blue, font: { bold: true, color: white, size: 11 }, verticalAlignment: "center" };
const headerStyle = { fill: lightBlue, font: { bold: true, color: dark }, verticalAlignment: "center", wrapText: true, borders: { preset: "inside", style: "thin", color: "#A6A6A6" } };

// Planning Rules
rules.getRange("A1:F1").merge();
rules.getRange("A1").values = [["KBMS Sprint Planning Rules"]];
rules.getRange("A1:F1").format = titleStyle;
rules.getRange("A3:C3").values = [["Priority", "Rank", "Default queue"]];
rules.getRange("A3:C3").format = headerStyle;
rules.getRange("A4:C8").values = [
  ["Very High", 1, "Sprint 1-3"], ["High", 2, "Sprint 3-5"], ["Medium", 3, "Sprint 5-6"], ["Low", 4, "Sprint 6-7"], ["Unspecified", 5, "Backlog - Set Priority"],
];
rules.getRange("E3:F3").values = [["Remark signal", "Planning use"]];
rules.getRange("E3:F3").format = headerStyle;
rules.getRange("E4:F9").values = [
  ["รีบ / High!!!!", "จัดก่อนภายใน Priority เดียวกัน"],
  ["ดำเนินการ", "จัดต้น Sprint เพื่อทำต่อเนื่อง"],
  ["Confirm / ไม่แน่ใจ / พิจารณา", "แยก TBC จนกว่าจะตัดสินใจ"],
  ["Investigate / ยาก", "วางท้ายกลุ่มและทำ Technical Spike ก่อน"],
  ["งานเกี่ยวข้องกัน", "รวมใน Sprint เดียวกันเพื่อลด Context switching"],
  ["Done / Cancelled", "ไม่นำมาวางแผน Sprint"],
];
rules.getRange("A11:F11").merge();
rules.getRange("A11").values = [["Capacity model: Sprint 2 สัปดาห์, ผู้ทำงาน 1 คนใช้ Codex ช่วยพัฒนา, คนเดิม Review/Verify และ Regression Test; จำกัดที่ 8 Capacity Points ต่อ Sprint และ Estimate รวมเวลาตรวจสอบแล้ว"]];
rules.getRange("A11:F11").format = { fill: paleYellow, font: { italic: true, color: "#7F6000" }, wrapText: true };
rules.getRange("A13:F13").merge(); rules.getRange("A13").values = [["Capacity assumptions & sizing guide"]]; rules.getRange("A13:F13").format = sectionStyle;
rules.getRange("A14:B18").values = [
  ["Sprint length", "2 weeks"], ["Gross working days", 10], ["Planned capacity", 8], ["Implementation", "Codex-assisted"], ["Verification", "Same person; included in points"],
];
rules.getRange("D14:F18").values = [
  ["Points", "Size", "Meaning"], [1, "XS", "Reuse/very small change"], [2, "S", "Small change + focused verify"], [3, "M", "Medium change + regression"], [5, "L", "Broad/uncertain change + wide regression"],
];
rules.getRange("D14:F14").format = headerStyle;
rules.getRange("A14:A18").format.font = { bold: true, color: navy };
rules.getRange("A:F").format.font = { name: "Tahoma", size: 10 };
rules.getRange("A:A").format.columnWidth = 20; rules.getRange("B:B").format.columnWidth = 24; rules.getRange("C:C").format.columnWidth = 24;
rules.getRange("D:D").format.columnWidth = 10; rules.getRange("E:E").format.columnWidth = 18; rules.getRange("F:F").format.columnWidth = 50;
rules.getRange("A1:F18").format.wrapText = true; rules.getRange("A1:F18").format.autofitRows();
rules.freezePanes.freezeRows(3);

// All analysis: every Issue/CR/Adjust row, including history.
const allHeaders = ["#ISSUE", "Type", "Page", "Detail", "Priority", "Created date", "Created by", "Status", "Resolved date", "Action Log/Remark", "Remark signal", "Planning disposition", "Priority rank"];
all.getRange("A1:M1").values = [allHeaders];
all.getRange("A1:M1").format = headerStyle;
const allData = targetRows.map(r => [r[0], upper(r[1]), r[2], r[3], r[4] || "Unspecified", excelDate(r[5]), r[6], r[7], excelDate(r[8]), r[9] || r[10] || "", signal(r[9] || r[10])]);
all.getRangeByIndexes(1, 0, allData.length, 11).values = allData;
for (let i = 0; i < allData.length; i++) {
  const excelRow = i + 2;
  all.getRange(`L${excelRow}`).formulas = [[`=IF(OR(H${excelRow}="Done",H${excelRow}="Cancelled"),"Excluded - "&H${excelRow},IF(H${excelRow}="TBC","TBC / Confirm",IFERROR(VLOOKUP(E${excelRow},'Planning Rules'!$A$4:$C$8,3,FALSE),"Backlog - Set Priority")))`]];
  all.getRange(`M${excelRow}`).formulas = [[`=IFERROR(VLOOKUP(E${excelRow},'Planning Rules'!$A$4:$C$8,2,FALSE),5)`]];
}
all.tables.add(`A1:M${allData.length + 1}`, true, "AllAnalysisTable").style = "TableStyleMedium2";
all.getRange("A:M").format.font = { name: "Tahoma", size: 9 };
all.getRange("A:A").format.columnWidth = 9; all.getRange("B:B").format.columnWidth = 11; all.getRange("C:C").format.columnWidth = 20;
all.getRange("D:D").format.columnWidth = 58; all.getRange("E:E").format.columnWidth = 13; all.getRange("F:F").format.columnWidth = 14;
all.getRange("G:G").format.columnWidth = 12; all.getRange("H:H").format.columnWidth = 13; all.getRange("I:I").format.columnWidth = 14;
all.getRange("J:J").format.columnWidth = 55; all.getRange("K:K").format.columnWidth = 22; all.getRange("L:L").format.columnWidth = 24; all.getRange("M:M").format.columnWidth = 12;
all.getRange(`C2:M${allData.length + 1}`).format.wrapText = true;
all.getRange(`A2:M${allData.length + 1}`).format.verticalAlignment = "top";
all.getRange(`F2:F${allData.length + 1}`).format.numberFormat = "dd-mmm-yy";
all.getRange(`I2:I${allData.length + 1}`).format.numberFormat = "dd-mmm-yy";
all.freezePanes.freezeRows(1); all.freezePanes.freezeColumns(2);

// Sprint Plan
plan.getRange("A1:N1").merge();
plan.getRange("A1").values = [["KBMS Proposed Sprint Plan — Open Issue / CR / Adjust"]];
plan.getRange("A1:N1").format = titleStyle;
plan.getRange("A2:N2").merge();
plan.getRange("A2").values = [["Capacity Model: 2-week Sprint | 1 คน + Codex | คนเดิม Verify | ไม่เกิน 8 Capacity Points/Sprint | ไม่รวม Done, Cancelled และ TBC"]];
plan.getRange("A2:N2").format = { fill: paleYellow, font: { italic: true, color: "#7F6000" } };
const planHeaders = ["Sprint", "Order", "#ISSUE", "Type", "Page", "Detail", "Priority", "Capacity points", "Created date", "Created by", "Status", "Action Log/Remark", "Remark signal", "Planning rationale"];
plan.getRange("A4:N4").values = [planHeaders]; plan.getRange("A4:N4").format = headerStyle;
const planned = targetRows.filter(r => sprintMap.has(Number(r[0]))).map(r => {
  const [s, order, points, rationale] = sprintMap.get(Number(r[0]));
  return [s, order, r[0], upper(r[1]), r[2], r[3], r[4], points, excelDate(r[5]), r[6], r[7] || "Open", r[9] || r[10] || "", signal(r[9] || r[10]), rationale];
}).sort((a,b) => Number(a[0].split(" ")[1]) - Number(b[0].split(" ")[1]) || a[1] - b[1]);
plan.getRangeByIndexes(4, 0, planned.length, 14).values = planned;
plan.tables.add(`A4:N${planned.length + 4}`, true, "SprintPlanTable").style = "TableStyleMedium2";
plan.getRange("A:N").format.font = { name: "Tahoma", size: 9 };
plan.getRange("A:A").format.columnWidth = 13; plan.getRange("B:B").format.columnWidth = 8; plan.getRange("C:C").format.columnWidth = 9;
plan.getRange("D:D").format.columnWidth = 10; plan.getRange("E:E").format.columnWidth = 20; plan.getRange("F:F").format.columnWidth = 60;
plan.getRange("G:G").format.columnWidth = 13; plan.getRange("H:H").format.columnWidth = 14; plan.getRange("I:I").format.columnWidth = 14; plan.getRange("J:J").format.columnWidth = 12;
plan.getRange("K:K").format.columnWidth = 11; plan.getRange("L:L").format.columnWidth = 50; plan.getRange("M:M").format.columnWidth = 22; plan.getRange("N:N").format.columnWidth = 52;
plan.getRange(`E5:N${planned.length + 4}`).format.wrapText = true;
plan.getRange(`A5:N${planned.length + 4}`).format.verticalAlignment = "top";
plan.getRange(`H5:H${planned.length + 4}`).format.numberFormat = "0";
plan.getRange(`I5:I${planned.length + 4}`).format.numberFormat = "dd-mmm-yy";
plan.freezePanes.freezeRows(4); plan.freezePanes.freezeColumns(3);

// TBC
tbc.getRange("A1:I1").merge(); tbc.getRange("A1").values = [["TBC & Decisions Required"]]; tbc.getRange("A1:I1").format = titleStyle;
tbc.getRange("A3:I3").values = [["#ISSUE", "Type", "Page", "Detail", "Priority", "Status", "Remark", "Decision needed", "Next action"]]; tbc.getRange("A3:I3").format = headerStyle;
const tbcRows = targetRows.filter(r => upper(r[7]) === "TBC").map(r => [r[0], upper(r[1]), r[2], r[3], r[4] || "Unspecified", r[7], r[9] || r[10] || "",
  Number(r[0]) === 43 ? "ยืนยันหลักเกณฑ์แยก PO OVER จากข้อมูล SAP" : "ยืนยันว่าจะใช้ Actual PO Date หรือ Delivery Date",
  "Owner นัดยืนยัน Requirement และกำหนด Priority ก่อนนำเข้า Sprint"]);
tbc.getRangeByIndexes(3, 0, tbcRows.length, 9).values = tbcRows;
tbc.tables.add(`A3:I${tbcRows.length + 3}`, true, "TBCDecisionTable").style = "TableStyleMedium4";
tbc.getRange("A:I").format.font = { name: "Tahoma", size: 10 };
const tbcCols = ["A","B","C","D","E","F","G","H","I"];
[10,10,18,54,13,12,52,45,48].forEach((w,i) => tbc.getRange(`${tbcCols[i]}:${tbcCols[i]}`).format.columnWidth = w);
tbc.getRange(`C4:I${tbcRows.length + 3}`).format.wrapText = true; tbc.getRange(`A4:I${tbcRows.length + 3}`).format.verticalAlignment = "top";
tbc.freezePanes.freezeRows(3);

// Summary
summary.getRange("A1:H1").merge(); summary.getRange("A1").values = [["KBMS Sprint Planning Summary"]]; summary.getRange("A1:H1").format = titleStyle;
summary.getRange("A2:H2").merge(); summary.getRange("A2").values = [["Scope: Type = ISSUE, CR, ADJUST | Source: Production_KBMS_issue log_202511203 (2).xlsx"]];
summary.getRange("A2:H2").format = { fill: lightBlue, font: { italic: true, color: navy } };
summary.getRange("A4:E4").values = [["Planned open items", "TBC decisions", "Execution model", "Verification", "Capacity / Sprint"]]; summary.getRange("A4:E4").format = sectionStyle;
summary.getRange("A5").formulas = [[`=COUNTA('Sprint Plan'!$C$5:$C$${planned.length + 4})`]];
summary.getRange("B5").formulas = [[`=COUNTA('TBC & Decisions'!$A$4:$A$${tbcRows.length + 3})`]];
summary.getRange("C5:D5").values = [["1 คน + Codex", "คนเดิมตรวจซ้ำ"]];
summary.getRange("E5").formulas = [["='Planning Rules'!B16"]];
summary.getRange("A5:E5").format = { fill: "#F2F2F2", font: { bold: true, color: navy, size: 16 }, horizontalAlignment: "center" };

summary.getRange("A8:H8").values = [["Sprint", "Primary priority", "Items", "Planned points", "Capacity", "Utilization", "Work packages / focus", "Planning note"]]; summary.getRange("A8:H8").format = headerStyle;
const sprintInfo = [
  ["Sprint 1", "Very High", "Upload >300 rows + Approve OTB blocker", "เต็ม Capacity; ปิดและ Verify #65 ก่อนเริ่ม #68"],
  ["Sprint 2", "Very High", "Export TXN calculation + Create Switch warnings", "เต็ม Capacity; #64 ใช้ Test case ร่วมกับ #58"],
  ["Sprint 3", "Very High / High", "SAP preview + OTB Remaining/Export Summary", "เต็ม Capacity; ตรวจยอดรวมและ Regression Report"],
  ["Sprint 4", "High", "Multi-filter across Draft PO / OTB pages", "4 Issue IDs แต่เป็น Feature package เดียวและ Reuse logic"],
  ["Sprint 5", "High / Medium", "Draft PO preview + Bulk Upload + Auto Segment", "เต็ม Capacity; ทำ #38 ก่อน แล้วต่อ #52/#8"],
  ["Sprint 6", "Medium / Low", "Report numeric types + Freeze header", "เหลือ Buffer 1 point สำหรับ Regression กว้างของ #49"],
  ["Sprint 7", "Low", "UI colors + Copy/paste Technical Spike", "ใช้ 6/8 points; เหลือ Buffer รองรับความไม่แน่นอนของ #3"],
];
for (let i = 0; i < sprintInfo.length; i++) {
  const row = 9 + i;
  const [sp, pri, focus, note] = sprintInfo[i];
  summary.getRange(`A${row}:B${row}`).values = [[sp, pri]];
  summary.getRange(`C${row}`).formulas = [[`=COUNTIF('Sprint Plan'!$A$5:$A$${planned.length + 4},A${row})`]];
  summary.getRange(`D${row}`).formulas = [[`=SUMIF('Sprint Plan'!$A$5:$A$${planned.length + 4},A${row},'Sprint Plan'!$H$5:$H$${planned.length + 4})`]];
  summary.getRange(`E${row}`).formulas = [["='Planning Rules'!$B$16"]];
  summary.getRange(`F${row}`).formulas = [[`=D${row}/E${row}`]];
  summary.getRange(`G${row}:H${row}`).values = [[focus, note]];
}
summary.getRange("A17:H17").merge(); summary.getRange("A17").values = [["ข้อควรระวังก่อน Commit Sprint"]]; summary.getRange("A17:H17").format = sectionStyle;
summary.getRange("A18:H21").merge(true);
summary.getRange("A18").values = [["1) 8 Capacity Points/Sprint เป็น Baseline เบื้องต้นสำหรับ 2 สัปดาห์ ต้องปรับจาก Velocity จริงหลังจบ 1-2 Sprint"]];
summary.getRange("A19").values = [["2) Estimate รวม Coding ด้วย Codex + Human Review/Verify แล้ว แต่ยังต้อง Refinement ก่อน Commit"]];
summary.getRange("A20").values = [["3) #16 และ #17 ควรตรวจว่าเป็น CR เดิมจาก UAT หรือไม่ เพื่อลดงานซ้ำ"]];
summary.getRange("A21").values = [["4) #43 และ #67 ต้องยืนยัน Requirement และ Priority ก่อนนำเข้า Sprint"]];
summary.getRange("A18:H21").format = { fill: paleYellow, font: { color: "#7F6000" }, wrapText: true };
summary.getRange("A:H").format.font = { name: "Tahoma", size: 10 };
const summaryCols = ["A","B","C","D","E","F","G","H"];
[22,20,18,18,18,14,52,52].forEach((w,i) => summary.getRange(`${summaryCols[i]}:${summaryCols[i]}`).format.columnWidth = w);
summary.getRange("F9:F15").format.numberFormat = "0%";
summary.getRange("A8:H21").format.wrapText = true; summary.getRange("A8:H21").format.verticalAlignment = "top"; summary.getRange("A8:H21").format.autofitRows();
summary.freezePanes.freezeRows(2);

// Priority / Sprint visual cues.
for (const sh of [plan, all]) {
  const used = sh.getUsedRange();
  if (used) used.format.borders = { preset: "inside", style: "thin", color: "#D9E1F2" };
}
plan.getRange(`G5:G${planned.length + 4}`).conditionalFormats.add("containsText", { text: "Very High", format: { fill: paleRed, font: { bold: true, color: "#9C0006" } } });
plan.getRange(`G5:G${planned.length + 4}`).conditionalFormats.add("containsText", { text: "High", format: { fill: paleOrange, font: { bold: true, color: "#9C5700" } } });
plan.getRange(`G5:G${planned.length + 4}`).conditionalFormats.add("containsText", { text: "Medium", format: { fill: paleYellow, font: { color: "#7F6000" } } });
plan.getRange(`G5:G${planned.length + 4}`).conditionalFormats.add("containsText", { text: "Low", format: { fill: paleGreen, font: { color: "#375623" } } });

await fs.mkdir(outputDir, { recursive: true });

// Compact verification before export.
console.log((await wb.inspect({ kind: "table", range: "'Sprint Summary'!A1:H21", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 10, maxChars: 10000 })).ndjson);
console.log((await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 200 }, summary: "formula error scan" })).ndjson);

for (const [name, range, file] of [
  ["Sprint Summary", "A1:H21", "final_summary.png"],
  ["Sprint Plan", `A1:N${planned.length + 4}`, "final_plan.png"],
  ["TBC & Decisions", `A1:I${tbcRows.length + 3}`, "final_tbc.png"],
  ["All Analysis", `A1:M${allData.length + 1}`, "final_all.png"],
  ["Planning Rules", "A1:F19", "final_rules.png"],
]) {
  const preview = await wb.render({ sheetName: name, range, scale: name === "Sprint Plan" || name === "All Analysis" ? 0.65 : 1, format: "png" });
  await fs.writeFile(`${outputDir}/${file}`, new Uint8Array(await preview.arrayBuffer()));
}

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(`OUTPUT ${outputPath}`);
