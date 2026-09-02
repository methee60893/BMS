import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath = "C:/Users/60893/Downloads/Production_KBMS_issue log_202511203.xlsx";
const builderDir = path.dirname(fileURLToPath(import.meta.url));
const workspaceRoot = path.resolve(builderDir, "../..");
const outputDir = path.resolve(process.cwd(), "../../outputs/kbms_issue_summary_20260624");
const outputPath = `${outputDir}/KBMS_issue_status_summary_20260624.xlsx`;

const evidence = {
  poUpload: "Handler/POUploadHandler.ashx.vb:57,110,137",
  poValidate: "class/POValidate.vb:86,342",
  matchingSp: "database/03_create_stored_procedures.sql:977-1092",
  dataOtb: "Handler/DataOTBHandler.ashx.vb:1521,1965,2095",
  remaining: "Handler/DataOTBRemainingHandler.ashx.vb:212",
  switchUpload: "Handler/SwitchUploadHandler.ashx.vb:201,264,422",
  validateSwitch: "Handler/ValidateHandler.ashx.vb:322-411",
  rounding: "class/share_class.vb:189-208",
  poRemarkMissing: "Handler/POUploadHandler.ashx.vb:274,313",
  switchResultUi: "createOTBswitching.aspx:604-614",
  filters: "draftPO.aspx:489-519; otbRemaining.aspx:356-392; approvedOTB.aspx:348-391",
};

const assessment = {
  1: ["แก้แล้วใน code", "API/matching มี logic แยกสถานะและ business key แล้ว", "UAT sync Actual PO และตรวจว่า GPO/Non GPO ไม่เข้าผิด flow", evidence.matchingSp],
  2: ["ยังไม่แก้", "ไม่พบ logic รองรับ copy/paste Amount THB ในหน้า Create Switch", "แก้ UI input/paste handler และทดสอบ paste จาก Excel", ""],
  3: ["ยังไม่แก้", "ไม่พบ fix ชัดเจนสำหรับ copy Amount CCY แล้ว Amount THB ผิดเมื่อ rate = 1", "ตรวจ JS calculation ใน Create Draft PO และเพิ่ม test paste/calculate", ""],
  4: ["แก้แล้วใน code", "Action By/Status By ถูกเก็บและแสดงตามสถานะแล้ว", "UAT รายงาน Draft/Actual status history", "Handler/DataPOHandler.ashx.vb:546,671"],
  5: ["ยังไม่แก้", "ไม่พบ implementation ชัดเจนเรื่อง fit page/freeze header ใน Switch TXN", "ปรับ DataTable/fixed header และทดสอบ viewport", ""],
  6: ["ยังไม่แก้", "เป็น cosmetic color shade; ไม่พบหลักฐาน fix เฉพาะใน code", "ยืนยัน mockup สีและปรับ CSS", ""],
  7: ["แก้แล้วใน code", "Draft PO upload/manual ใช้ duplicate validation แล้ว", "UAT duplicate Draft PO no. ทั้ง manual และ upload", evidence.poValidate],
  8: ["ยังไม่แก้", "ไม่พบ auto-fill Segment จาก Brand/Vendor", "เพิ่ม mapping rule brand/vendor -> segment", ""],
  9: ["ยังไม่แก้", "หน้า Draft PO filter ยังเป็น single-value select", "ปรับ UI เป็น multi-select และแก้ query ให้รับหลายค่า", evidence.filters],
  10: ["ยังไม่แก้", "Draft PO มี column date แต่ไม่พบ date filter", "เพิ่ม from/to Draft PO date filter ใน UI และ handler", "draftPO.aspx:199,489-519"],
  11: ["แก้แล้วใน code", "Edit Draft PO update field/status/remark แล้ว", "UAT edit/save แล้ว reload data", "Handler/DataPOHandler.ashx.vb:321-355"],
  12: ["แก้แล้วใน code", "OTB Remaining มี Total Actual + Draft PO แล้ว", "UAT remaining/export", evidence.remaining],
  13: ["แก้แล้วใน code", "Upload Draft PO re-validate ก่อน save และใช้ transaction ไม่ควร partial save", "UAT mixed pass/fail upload", evidence.poUpload],
  14: ["ยังไม่แก้", "ไม่พบ filter by color / color-change logic ใน Match Actual PO", "กำหนดสี/สถานะและเพิ่ม filter", ""],
  15: ["ยังไม่แก้", "OTB Remaining filter ยังเป็น single-value select", "ปรับ multi-select และ SQL WHERE IN", evidence.filters],
  16: ["ยังไม่แก้", "Draft OTB filter ยังเป็น single-value select", "ปรับ multi-select และ handler", evidence.filters],
  17: ["ยังไม่แก้", "Approved OTB filter ยังเป็น single-value select", "ปรับ multi-select และ handler", evidence.filters],
  18: ["แก้แล้วใน code", "Actual report/export มี Company field แล้ว", "UAT export Actual PO report", "Handler/DataPOHandler.ashx.vb:585-610"],
  19: ["แก้แล้วใน code", "Confirm match update Draft/Actual เป็น Matched พร้อม status date/by", "UAT confirm match", "Handler/POMatchingHandler.ashx.vb:475,510"],
  20: ["แก้แล้วใน code", "มี flow bulk upload switch OTB แล้ว", "UAT bulk switch upload กับ SAP response", evidence.switchUpload],
  21: ["แก้แล้วใน code", "Final match table บังคับ Actual/Draft unique กัน Draft เดียว match หลาย Actual", "UAT PO ที่เคย duplicate", evidence.matchingSp],
  22: ["แก้แล้วใน code", "Auto-match rank หา best pair และกันคู่ที่ถูกใช้แล้ว", "UAT sync/rematch case เดิม", evidence.matchingSp],
  23: ["แก้แล้วใน code", "Edit เฉพาะ Draft PO no. ไม่ trigger budget validation ที่ไม่จำเป็น", "UAT edit PO no only", "Handler/DataPOHandler.ashx.vb:241-277"],
  24: ["แก้แล้วใน code", "Sync Actual PO มี logic handle moved/stale/deletion flag", "ตรวจ production data case Nov 2025", "Handler/POMatchingHandler.ashx.vb:788-861"],
  25: ["แก้แล้วใน code", "Preview/save ใช้ ValidateBatch ซ้ำ ลดความต่างระหว่าง preview กับ submit", "UAT upload pass/fail mixed", evidence.poUpload],
  26: ["แก้บางส่วน", "เอกสารระบุ Done แต่ code Draft PO upload ยัง set RowIndex = i + 1 ไม่ใช่เลขแถว Excel จริง", "แก้เป็น Excel row number และ UAT ร่วมกับ #57", evidence.poRemarkMissing],
  27: ["แก้แล้วใน code", "Auto-match/sync ป้องกัน one Draft match หลาย Actual; rollback manual แยกไปอยู่ #39", "UAT rematch case KPC-LM-SP-12/25-443", evidence.matchingSp],
  28: ["รอ deploy/UAT", "Issue log เป็น Re-open; code มี logic sync/moved row แต่ต้องยืนยัน deploy production", "Deploy DB/handler แล้วเทียบ SAP vs KBMS รายวัน", "Handler/POMatchingHandler.ashx.vb:917-929"],
  29: ["แก้แล้วใน code", "Actual/Draft update Status_Date แล้ว", "UAT matched date ในหน้าและ export", "database/03_create_stored_procedures.sql:42,1065,1082"],
  30: ["แก้แล้วใน code", "Sync summary ยกเลิก duplicate actual ด้วย row_number", "UAT duplicate Actual PO same dimension", "database/03_create_stored_procedures.sql:714-767"],
  31: ["แก้แล้วใน code", "Export summary เพิ่ม Actual PO/Draft PO/Total Actual + Draft PO", "UAT export summary ตามตัวอย่าง", evidence.dataOtb],
  32: ["แก้แล้วใน code", "Auto-match ไม่ reuse Actual/Draft ที่ถูก match แล้ว", "UAT duplicate matched actual", evidence.matchingSp],
  33: ["แก้แล้วใน code", "Unique DraftPO_ID ใน final matches ป้องกัน Draft เดียวไปหลาย Actual", "UAT case อ้างอิง #27", evidence.matchingSp],
  34: ["ยกเลิกตาม issue log", "เอกสารยกเลิก; logic ปัจจุบัน cancel/ไม่แสดง actual amount = 0", "ไม่ต้องทำต่อ ยกเว้น business เปลี่ยน requirement", "database/03_create_stored_procedures.sql:835-849"],
  35: ["ยกเลิกตาม issue log", "เอกสารยกเลิกและโยงไปทำ #58 แทน", "ตามต่อที่ #58", ""],
  36: ["ยกเลิกตาม issue log", "เอกสารยกเลิก/เกี่ยวกับ #35", "ตามต่อที่ #58 หากยังต้อง warning month", ""],
  37: ["แก้แล้วใน code", "Report คำนวณ switch out/balance out/carry out เป็นลบใน total", "UAT Export TXN Out categories", "Handler/DataOTBHandler.ashx.vb:38-60,1965"],
  38: ["ยังไม่แก้", "Preview budget insufficient ยังไม่เห็น cate/brand context ตามที่ขอครบ", "เพิ่ม column context ใน error preview", ""],
  39: ["ยังไม่แก้", "ไม่พบ UI/function rollback matched Actual/Draft แบบ user action", "ออกแบบ rollback/unmatch flow พร้อม audit log", ""],
  40: ["แก้แล้วใน code", "Matching ใช้ business key รวม category/segment/brand/vendor ไม่ใช่ PO no อย่างเดียว", "UAT sheet ISSUE#40", evidence.matchingSp],
  41: ["แก้บางส่วน", "รองรับ moved/stale row บางกรณี แต่ยังไม่ใช่ auto detect เปลี่ยน PO number/cate/amount แบบทั่วไป", "กำหนด rule detection จาก SAP cancellation/change แล้วเพิ่ม logic", "Handler/POMatchingHandler.ashx.vb:788-861"],
  42: ["แก้แล้วใน code", "Auto-match update Draft status เป็น Matching และ clear/release stale reference", "รัน data repair/UAT remaining negative", evidence.matchingSp],
  43: ["ยังไม่แก้", "Actual PO report ใช้ Actual_PO_Summary.Remark ไม่ได้ propagate Draft remark ชัดเจน", "Join/copy Draft PO remark ไป Actual report ตาม rule", "database/03_create_stored_procedures.sql:43"],
  44: ["แก้บางส่วน", "Auto-match fuzzy amount รองรับภายใน tolerance 10%; ถ้าต่างมากกว่านั้นยังไม่ match", "ยืนยัน tolerance business และปรับ matching rule", "database/03_create_stored_procedures.sql:1009-1058"],
  45: ["แก้แล้วใน code", "Manual/bulk switch มี available budget validation", "UAT switch over budget แล้วต้อง block", `${evidence.validateSwitch}; ${evidence.switchUpload}`],
  46: ["แก้แล้วใน code", "Draft PO upload ใช้ aggregate budget validation ก่อน save", "UAT upload over budget", evidence.poValidate],
  47: ["แก้แล้วใน code", "Bulk switch upload เช็ค remaining และ used in batch", "UAT brand CHY case", evidence.switchUpload],
  48: ["แก้แล้วใน code", "Manual match เก็บ Draft_PO_ID_Ref และ sync preserve matched/force matching", "UAT manual match แล้ว sync SAP", "Handler/POMatchingHandler.ashx.vb:148,158"],
  49: ["แก้บางส่วน", "บาง export format Decimal เป็น number แล้ว แต่ยังไม่ยืนยันทุก report field", "ไล่ทุก export/report และกำหนด numeric format ทั้งหมด", "Handler/DataOTBHandler.ashx.vb:344-365"],
  50: ["ยังไม่แก้", "Draft PO upload template remark ไม่ถูก save; BuildDraftPOBulkTable set Remark = DBNull", "เพิ่ม Remark ใน FrontendPORow/preview/save/bulk table", evidence.poRemarkMissing],
  51: ["ยังไม่แก้", "Actual PO remark ยังไม่ดึง Draft PO remark เช่น PO OVER", "ปรับ SP/report ให้แสดง Draft remark เมื่อ Actual matched", "database/03_create_stored_procedures.sql:43"],
  52: ["แก้บางส่วน", "Backend มีชื่อ category/segment/brand แต่ upload result UI ยังโชว์หลัก ๆ แค่ company/vendor", "ปรับ renderBulkResults ให้แสดงครบทุก dimension", evidence.switchResultUi],
  53: ["แก้แล้วใน code", "เหมือน #46: upload Draft PO over budget ถูก validate ก่อน save", "UAT upload over budget case เดิม", evidence.poValidate],
  54: ["แก้แล้วใน code", "Sync handles PO moved month/stale rows", "UAT PO2011009169 Feb/Mar", "Handler/POMatchingHandler.ashx.vb:788-861"],
  55: ["แก้แล้วใน code", "Matched Draft/Actual update amount/status/ref และ remaining ใช้ usage calc ใหม่", "UAT matched amount zero/remaining", `${evidence.matchingSp}; ${evidence.remaining}`],
  56: ["ยังไม่แก้", "Manual match ยังหา Draft pair ตาม key; ยังไม่รองรับ no-pair free matching", "เพิ่ม manual selection/search และ validation rule", "Handler/POMatchingHandler.ashx.vb:81-117"],
  57: ["ยังไม่แก้", "Draft PO upload row no ยังเป็น i + 1 ไม่ใช่ Excel row sequence จริง", "แก้ RowIndex เป็น Excel row และ UAT", evidence.poRemarkMissing],
  58: ["ยังไม่แก้", "ไม่พบข้อความ/สี warning เดือนไม่สอดคล้องกัน", "เพิ่ม month mismatch validation + yellow warning ใน preview", ""],
  59: ["แก้บางส่วน", "หน้า Switch TXN แสดง seconds แล้ว แต่ export/บาง formatter ยังใช้ HH:mm", "ปรับ export/ทุก formatter ให้ใช้ HH:mm:ss ถ้าต้องการสอดคล้อง", "Handler/DataOTBHandler.ashx.vb:194,713"],
  60: ["ยังไม่แก้", "ไม่พบ logic ดึง PO rate จาก SAP สำหรับ Create Draft PO", "เพิ่ม SAP rate lookup หรือยืนยันว่า user upload rate เอง", ""],
  61: ["ยังไม่แก้", "ข้อความ over budget ยัง generic และบาง preview context ไม่ครบ", "ปรับ error message ให้ระบุ brand/category/remaining ชัดเจน", "class/POValidate.vb:342"],
  62: ["ยังไม่แก้", "ไม่พบ auto carry balance ใน code ที่ตรวจ", "ออกแบบ carry balance job/rule และ audit", ""],
  63: ["แก้แล้วใน code", "Amount ถูก round/format ก่อนส่ง SAP", "UAT decimal > 3 ตำแหน่ง", evidence.rounding],
  64: ["ยังไม่แก้", "ไม่พบ decimal highlight yellow ใน switch budget preview", "เพิ่ม decimal detection/highlight ตาม requirement", ""],
  65: ["ต้องตรวจ production data", "ไม่พบ hard limit 300 rows ใน code; มี BulkCopy และ timeout 300 วินาที จึงอาจเป็น timeout/request/SAP payload", "ทดสอบ upload >300 rows บน production-like env และดู log timeout/API", "Handler/UploadHandler.ashx.vb:540,803; Handler/POUploadHandler.ashx.vb:137; Web.config:35"],
};

const statusRank = {
  "ยังไม่แก้": 1,
  "แก้บางส่วน": 2,
  "รอ deploy/UAT": 3,
  "ต้องตรวจ production data": 4,
  "แก้แล้วใน code": 8,
  "ยกเลิกตาม issue log": 9,
};

const priorityRank = {
  "Very High": 1,
  "High": 2,
  "Medium": 3,
  "Low": 4,
  "blank": 5,
  "": 5,
};

function cellToText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value;
  return value;
}

function excelSerialToDate(n) {
  if (typeof n !== "number" || n < 20000 || n > 80000) return n;
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + Math.round(n) * 86400000);
}

function normalizeDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number") return excelSerialToDate(value);
  return value || "";
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function priorityToRank(priority) {
  return priorityRank[priority || "blank"] || 5;
}

function detailSortRank(row) {
  return (statusRank[row.systemStatus] || 5) * 100 + priorityToRank(row.priority) * 10 + row.issueNo / 100;
}

function statusFill(status) {
  if (status === "ยังไม่แก้") return "#FEE2E2";
  if (status === "แก้บางส่วน") return "#FEF3C7";
  if (status === "รอ deploy/UAT" || status === "ต้องตรวจ production data") return "#DBEAFE";
  if (status === "แก้แล้วใน code") return "#DCFCE7";
  if (status === "ยกเลิกตาม issue log") return "#E5E7EB";
  return "#FFFFFF";
}

const sourceBlob = await FileBlob.load(sourcePath);
const sourceWorkbook = await SpreadsheetFile.importXlsx(sourceBlob);
const sourceSheet = sourceWorkbook.worksheets.getItem("ISSUE");
const values = sourceSheet.getUsedRange().values;
const headers = values[0].map((h) => String(h || "").trim());
const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

const rows = values.slice(1)
  .filter((r) => r[idx["#ISSUE"]] !== null && r[idx["#ISSUE"]] !== undefined && r[idx["#ISSUE"]] !== "")
  .map((r) => {
    const issueNo = Number(r[idx["#ISSUE"]]);
    const rowAssessment = assessment[issueNo] || ["ต้องตรวจ production data", "ยังไม่มี assessment mapping", "ตรวจ code เพิ่ม", ""];
    const priority = String(r[idx["Priority"]] || "blank").trim() || "blank";
    const logStatus = String(r[idx["Status"]] || "Open/blank").trim() || "Open/blank";
    const detail = String(r[idx["Detail"]] || "").replace(/\s+/g, " ").trim();
    const page = String(r[idx["Page"]] || "").trim();
    const type = String(r[idx["Type"]] || "").trim();
    const summary = detail.length > 160 ? `${detail.slice(0, 157)}...` : detail;
    return {
      issueNo,
      priority,
      logStatus,
      page,
      type,
      summary,
      detail,
      systemStatus: rowAssessment[0],
      why: rowAssessment[1],
      nextAction: rowAssessment[2],
      evidence: rowAssessment[3],
      createdDate: normalizeDate(r[idx["Created date"]]),
      resolvedDate: normalizeDate(r[idx["Resolved date"]]),
      remark: String(r[idx["Action Log/Remark"]] || "").replace(/\s+/g, " ").trim(),
    };
  })
  .map((row) => ({ ...row, sortRank: detailSortRank(row) }))
  .sort((a, b) => a.sortRank - b.sortRank);

const openRows = rows.filter((r) => !["แก้แล้วใน code", "ยกเลิกตาม issue log"].includes(r.systemStatus));

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Dashboard");
const detailSheet = workbook.worksheets.add("Issue Detail");
const actionSheet = workbook.worksheets.add("Next Actions");
const notesSheet = workbook.worksheets.add("Source Notes");

for (const sheet of [dashboard, detailSheet, actionSheet, notesSheet]) {
  sheet.showGridLines = false;
}

const detailHeaders = [
  "Sort Rank", "Issue No", "Priority", "Issue Log Status", "System Check Status",
  "Page", "Type", "Summary", "Original Detail", "Why / Evidence",
  "Next Action", "Code Evidence", "Created Date", "Resolved Date", "Issue Remark"
];
const detailData = rows.map((r) => [
  r.sortRank, r.issueNo, r.priority, r.logStatus, r.systemStatus, r.page, r.type,
  r.summary, r.detail, r.why, r.nextAction, r.evidence, r.createdDate, r.resolvedDate, r.remark,
]);
detailSheet.getRangeByIndexes(0, 0, detailData.length + 1, detailHeaders.length).values = [detailHeaders, ...detailData];
const detailLastRow = detailData.length + 1;
const detailLastCol = colLetter(detailHeaders.length);
detailSheet.tables.add(`A1:${detailLastCol}${detailLastRow}`, true, "IssueDetailTable");
detailSheet.freezePanes.freezeRows(1);
detailSheet.freezePanes.freezeColumns(2);
detailSheet.getRange(`A1:${detailLastCol}1`).format = {
  fill: "#0F172A",
  font: { bold: true, color: "#FFFFFF" },
};
detailSheet.getRange(`A1:${detailLastCol}${detailLastRow}`).format.borders = { preset: "inside", style: "thin", color: "#E5E7EB" };
detailSheet.getRange(`A2:${detailLastCol}${detailLastRow}`).format.wrapText = true;
detailSheet.getRange(`A2:A${detailLastRow}`).format.numberFormat = "0.00";
detailSheet.getRange(`B2:B${detailLastRow}`).format.numberFormat = "0";
detailSheet.getRange(`M2:N${detailLastRow}`).format.numberFormat = "yyyy-mm-dd";
for (let i = 0; i < rows.length; i++) {
  const rowNum = i + 2;
  detailSheet.getRange(`A${rowNum}:${detailLastCol}${rowNum}`).format.fill = statusFill(rows[i].systemStatus);
}
const detailWidths = [10, 9, 12, 16, 20, 20, 14, 42, 58, 58, 50, 54, 14, 14, 56];
detailWidths.forEach((w, i) => {
  detailSheet.getRange(`${colLetter(i + 1)}:${colLetter(i + 1)}`).format.columnWidth = w;
});
detailSheet.getRange(`A1:${detailLastCol}${detailLastRow}`).format.autofitRows();

dashboard.getRange("A1:H1").merge();
dashboard.getRange("A1").values = [["KBMS Production Issue Summary"]];
dashboard.getRange("A1").format = {
  fill: "#0F172A",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};
dashboard.getRange("A2:H2").merge();
dashboard.getRange("A2").values = [[`Source: ${sourcePath} | Generated: 2026-06-24 | Check: issue log + current source code`]];
dashboard.getRange("A2").format = { fill: "#E5E7EB", font: { color: "#374151" } };

dashboard.getRange("A4:B8").values = [
  ["Issue Log Status", "Count"],
  ["Done", null],
  ["Open/blank", null],
  ["Re-open", null],
  ["Cancelled", null],
];
dashboard.getRange("B5:B8").formulas = [
  [`=COUNTIF('Issue Detail'!$D$2:$D$${detailLastRow},A5)`],
  [`=COUNTIF('Issue Detail'!$D$2:$D$${detailLastRow},A6)`],
  [`=COUNTIF('Issue Detail'!$D$2:$D$${detailLastRow},A7)`],
  [`=COUNTIF('Issue Detail'!$D$2:$D$${detailLastRow},A8)`],
];

dashboard.getRange("D4:E10").values = [
  ["System Check Status", "Count"],
  ["ยังไม่แก้", null],
  ["แก้บางส่วน", null],
  ["รอ deploy/UAT", null],
  ["ต้องตรวจ production data", null],
  ["แก้แล้วใน code", null],
  ["ยกเลิกตาม issue log", null],
];
dashboard.getRange("E5:E10").formulas = [
  [`=COUNTIF('Issue Detail'!$E$2:$E$${detailLastRow},D5)`],
  [`=COUNTIF('Issue Detail'!$E$2:$E$${detailLastRow},D6)`],
  [`=COUNTIF('Issue Detail'!$E$2:$E$${detailLastRow},D7)`],
  [`=COUNTIF('Issue Detail'!$E$2:$E$${detailLastRow},D8)`],
  [`=COUNTIF('Issue Detail'!$E$2:$E$${detailLastRow},D9)`],
  [`=COUNTIF('Issue Detail'!$E$2:$E$${detailLastRow},D10)`],
];

dashboard.getRange("G4:H9").values = [
  ["Priority", "Count"],
  ["Very High", null],
  ["High", null],
  ["Medium", null],
  ["Low", null],
  ["blank", null],
];
dashboard.getRange("H5:H9").formulas = [
  [`=COUNTIF('Issue Detail'!$C$2:$C$${detailLastRow},G5)`],
  [`=COUNTIF('Issue Detail'!$C$2:$C$${detailLastRow},G6)`],
  [`=COUNTIF('Issue Detail'!$C$2:$C$${detailLastRow},G7)`],
  [`=COUNTIF('Issue Detail'!$C$2:$C$${detailLastRow},G8)`],
  [`=COUNTIF('Issue Detail'!$C$2:$C$${detailLastRow},G9)`],
];

for (const range of ["A4:B8", "D4:E10", "G4:H9"]) {
  dashboard.getRange(range).format.borders = { preset: "all", style: "thin", color: "#CBD5E1" };
}
dashboard.getRange("A4:B4").format = { fill: "#1D4ED8", font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("D4:E4").format = { fill: "#1D4ED8", font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("G4:H4").format = { fill: "#1D4ED8", font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("A12:H12").merge();
dashboard.getRange("A12").values = [["Top Open Items To Follow Up"]];
dashboard.getRange("A12").format = { fill: "#334155", font: { bold: true, color: "#FFFFFF" } };
const topHeaders = ["Issue No", "Priority", "System Status", "Page", "Summary", "Next Action"];
const topRows = openRows.slice(0, 10).map((r) => [r.issueNo, r.priority, r.systemStatus, r.page, r.summary, r.nextAction]);
dashboard.getRangeByIndexes(12, 0, topRows.length + 1, topHeaders.length).values = [topHeaders, ...topRows];
dashboard.getRange(`A13:F${13 + topRows.length}`).format.borders = { preset: "all", style: "thin", color: "#CBD5E1" };
dashboard.getRange("A13:F13").format = { fill: "#E2E8F0", font: { bold: true, color: "#111827" } };
dashboard.getRange(`A14:F${13 + topRows.length}`).format.wrapText = true;
dashboard.freezePanes.freezeRows(3);
[14, 12, 20, 28, 60, 58, 14, 12].forEach((w, i) => {
  dashboard.getRange(`${colLetter(i + 1)}:${colLetter(i + 1)}`).format.columnWidth = w;
});

const actionHeaders = ["Issue No", "Priority", "System Check Status", "Page", "Type", "Summary", "Why / Evidence", "Next Action", "Code Evidence"];
const actionRows = openRows.map((r) => [r.issueNo, r.priority, r.systemStatus, r.page, r.type, r.summary, r.why, r.nextAction, r.evidence]);
actionSheet.getRangeByIndexes(0, 0, actionRows.length + 1, actionHeaders.length).values = [actionHeaders, ...actionRows];
const actionLastRow = actionRows.length + 1;
const actionLastCol = colLetter(actionHeaders.length);
actionSheet.tables.add(`A1:${actionLastCol}${actionLastRow}`, true, "NextActionsTable");
actionSheet.getRange(`A1:${actionLastCol}1`).format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
actionSheet.getRange(`A2:${actionLastCol}${actionLastRow}`).format.wrapText = true;
actionSheet.getRange(`A1:${actionLastCol}${actionLastRow}`).format.borders = { preset: "inside", style: "thin", color: "#E5E7EB" };
actionSheet.freezePanes.freezeRows(1);
actionSheet.freezePanes.freezeColumns(1);
for (let i = 0; i < openRows.length; i++) {
  const rowNum = i + 2;
  actionSheet.getRange(`A${rowNum}:${actionLastCol}${rowNum}`).format.fill = statusFill(openRows[i].systemStatus);
}
[10, 12, 22, 18, 14, 50, 55, 55, 55].forEach((w, i) => {
  actionSheet.getRange(`${colLetter(i + 1)}:${colLetter(i + 1)}`).format.columnWidth = w;
});
actionSheet.getRange(`A1:${actionLastCol}${actionLastRow}`).format.autofitRows();

notesSheet.getRange("A1:B9").values = [
  ["Field", "Value"],
  ["Source workbook", sourcePath],
  ["Source sheet", "ISSUE"],
  ["Generated date", "2026-06-24"],
  ["Method", "Read issue log and compare against current source code in D:/CIE/BMS"],
  ["Important caveat", "System Check Status means code-level evidence was found. Production status still depends on deployment, database scripts, SAP data, and UAT."],
  ["Workbook rows", rows.length],
  ["Open action rows", openRows.length],
  ["Recommended next step", "Deploy/confirm DB scripts, then UAT the open and partial issues in Next Actions."],
];
notesSheet.getRange("A1:B1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
notesSheet.getRange("A1:B9").format.borders = { preset: "all", style: "thin", color: "#CBD5E1" };
notesSheet.getRange("A:B").format.wrapText = true;
notesSheet.getRange("A:A").format.columnWidth = 24;
notesSheet.getRange("B:B").format.columnWidth = 110;

const statusColorRows = [
  ["Status", "Meaning"],
  ["ยังไม่แก้", "ไม่พบ implementation หรือยังไม่ตรง requirement"],
  ["แก้บางส่วน", "มี code รองรับบางส่วน แต่ยังมี gap"],
  ["รอ deploy/UAT", "มี code รองรับ แต่ issue log หรือหลักฐานบอกว่ายังต้อง deploy/UAT"],
  ["ต้องตรวจ production data", "ยังต้องทดสอบกับ environment/data จริง"],
  ["แก้แล้วใน code", "พบ logic รองรับใน source code ปัจจุบัน"],
  ["ยกเลิกตาม issue log", "issue ถูก cancel ตามเอกสาร"],
];
notesSheet.getRangeByIndexes(11, 0, statusColorRows.length, 2).values = statusColorRows;
notesSheet.getRange("A12:B12").format = { fill: "#334155", font: { bold: true, color: "#FFFFFF" } };
for (let i = 1; i < statusColorRows.length; i++) {
  notesSheet.getRange(`A${12 + i}:B${12 + i}`).format.fill = statusFill(statusColorRows[i][0]);
}
notesSheet.getRange(`A12:B${11 + statusColorRows.length}`).format.borders = { preset: "all", style: "thin", color: "#CBD5E1" };

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(formulaErrors.ndjson);

const dashboardInspect = await workbook.inspect({
  kind: "table",
  sheetId: "Dashboard",
  range: "A1:H24",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 8,
});
console.log(dashboardInspect.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const [sheetName, fileName] of [
  ["Dashboard", "preview_dashboard.png"],
  ["Issue Detail", "preview_issue_detail.png"],
  ["Next Actions", "preview_next_actions.png"],
  ["Source Notes", "preview_source_notes.png"],
]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, fileName), new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`Saved ${outputPath}`);
