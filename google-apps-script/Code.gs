const DEFAULT_SHEET_ID = "1yydaayto6uVSTJ8Qav84kBcZ_HDQiBGKBIkBOGjlBvU";
const SCREENING_SHEET = "NCD_SCREENING_RESULTS";
const ANALYSIS_SHEET = "NCD_AI_ANALYSIS";
const REFERRAL_SHEET = "NCD_REFERRAL_FOLLOWUP";

function setupNcdWorkbook() {
  const ss = SpreadsheetApp.openById(DEFAULT_SHEET_ID);
  setupWorkbook_(ss);
  return {
    ok: true,
    spreadsheetId: DEFAULT_SHEET_ID,
    sheets: [SCREENING_SHEET, ANALYSIS_SHEET, REFERRAL_SHEET],
  };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    const sheetId = payload.sheetId || DEFAULT_SHEET_ID;
    const record = payload.record || {};
    const ss = SpreadsheetApp.openById(sheetId);

    setupWorkbook_(ss);
    appendRecord_(ss.getSheetByName(SCREENING_SHEET), getScreeningColumns_(), record);
    appendRecord_(ss.getSheetByName(ANALYSIS_SHEET), getAnalysisColumns_(), buildAnalysisRecord_(record));

    if (record.riskLevel === "high") {
      appendRecord_(ss.getSheetByName(REFERRAL_SHEET), getReferralColumns_(), buildReferralRecord_(record));
    }

    return json_({
      ok: true,
      message: "saved",
      sheets: [SCREENING_SHEET, ANALYSIS_SHEET, REFERRAL_SHEET],
      savedAt: new Date().toISOString(),
    });
  } catch (error) {
    return json_({ ok: false, error: String(error && error.message ? error.message : error) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const sheetId = (e && e.parameter && e.parameter.sheetId) || DEFAULT_SHEET_ID;
  const ss = SpreadsheetApp.openById(sheetId);

  setupWorkbook_(ss);

  return json_({
    ok: true,
    service: "Khamyai NCD Screening",
    action: action || "setup",
    spreadsheetId: sheetId,
    sheets: [SCREENING_SHEET, ANALYSIS_SHEET, REFERRAL_SHEET],
  });
}

function setupWorkbook_(ss) {
  setupSheet_(ss, SCREENING_SHEET, getScreeningColumns_(), "ผลการคัดกรองสุขภาพ NCDs รายบุคคล");
  setupSheet_(ss, ANALYSIS_SHEET, getAnalysisColumns_(), "ผลวิเคราะห์และจัดกลุ่มความเสี่ยง");
  setupSheet_(ss, REFERRAL_SHEET, getReferralColumns_(), "รายการเสี่ยงสูงที่ควรส่งต่อ/ติดตาม");
}

function setupSheet_(ss, sheetName, columns, title) {
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const keys = columns.map((column) => column.key);
  const labels = columns.map((column) => column.label);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, labels.length).setValues([labels]);
    sheet.getRange(2, 1, 1, keys.length).setValues([keys]);
  } else {
    const currentKeys = sheet.getRange(2, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].filter(String);
    const missingColumns = columns.filter((column) => !currentKeys.includes(column.key));
    if (missingColumns.length) {
      const startColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, startColumn, 1, missingColumns.length).setValues([missingColumns.map((column) => column.label)]);
      sheet.getRange(2, startColumn, 1, missingColumns.length).setValues([missingColumns.map((column) => column.key)]);
    }
  }

  sheet.setFrozenRows(2);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold").setBackground("#087f7a").setFontColor("#ffffff");
  sheet.getRange(2, 1, 1, sheet.getLastColumn()).setFontWeight("bold").setBackground("#e8f2ef").setFontColor("#075f5c");
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  ensureFilter_(sheet);
  sheet.getRange("A1").setNote(title);
}

function appendRecord_(sheet, columns, record) {
  const keys = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = keys.map((key) => formatValue_(record[key]));
  sheet.appendRow(row);
}

function buildAnalysisRecord_(record) {
  return {
    recordId: record.recordId,
    createdAt: record.createdAt,
    screeningDate: record.screeningDate,
    fullName: record.fullName,
    personId: record.personId,
    age: record.age,
    gender: record.gender,
    village: record.village,
    houseNo: record.houseNo,
    volunteer: record.volunteer,
    createdByName: record.createdByName,
    bmi: record.bmi,
    bloodPressure: [record.sbp, record.dbp].filter(Boolean).join("/"),
    glucose: record.glucose,
    cholesterol: record.cholesterol,
    phq2: record.phq2,
    riskLevel: record.riskLevel,
    riskLabel: record.riskLabel,
    abnormalFlags: record.flags,
    aiAdvice: record.advice,
    followUp: record.followUp,
    referralRequired: record.riskLevel === "high" ? "ใช่" : "ไม่ใช่",
    monthKey: getMonthKey_(record.screeningDate || record.createdAt),
  };
}

function buildReferralRecord_(record) {
  return {
    recordId: record.recordId,
    createdAt: record.createdAt,
    screeningDate: record.screeningDate,
    fullName: record.fullName,
    personId: record.personId,
    village: record.village,
    houseNo: record.houseNo,
    volunteer: record.volunteer,
    phone: record.phone || "",
    riskLabel: record.riskLabel,
    abnormalFlags: record.flags,
    followUp: record.followUp,
    referralStatus: "รอส่งต่อ/ติดตาม",
    referralDate: "",
    serviceUnit: "รพ.สต./หน่วยบริการ",
    note: "",
  };
}

function ensureFilter_(sheet) {
  if (sheet.getFilter()) return;
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastColumn = sheet.getLastColumn();
  sheet.getRange(2, 1, lastRow - 1, lastColumn).createFilter();
}

function formatValue_(value) {
  if (value === undefined || value === null) return "";
  return Array.isArray(value) ? value.join(" | ") : value;
}

function getMonthKey_(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value).slice(0, 7);
  return Utilities.formatDate(date, "Asia/Bangkok", "yyyy-MM");
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getScreeningColumns_() {
  return [
    col_("recordId", "รหัสรายการ"),
    col_("createdAt", "เวลาบันทึก"),
    col_("createdByUserId", "รหัสผู้บันทึก"),
    col_("createdByName", "ผู้บันทึก"),
    col_("personId", "เลขบัตรประชาชน"),
    col_("fullName", "ชื่อ-สกุล"),
    col_("age", "อายุ"),
    col_("gender", "เพศ"),
    col_("village", "หมู่บ้าน"),
    col_("houseNo", "บ้านเลขที่"),
    col_("volunteer", "อสม.ที่รับผิดชอบ"),
    col_("disease", "โรคประจำตัว"),
    col_("screeningDate", "วันที่คัดกรอง"),
    col_("screener", "ผู้คัดกรอง"),
    col_("mode", "ช่องทางคัดกรอง"),
    col_("eligibility", "สิทธิ์/ข้อมูลพื้นฐาน"),
    col_("weight", "น้ำหนัก (กก.)"),
    col_("height", "ส่วนสูง (ซม.)"),
    col_("bmi", "BMI"),
    col_("waist", "รอบเอว (ซม.)"),
    col_("sbp", "ความดันตัวบน"),
    col_("dbp", "ความดันตัวล่าง"),
    col_("pulse", "ชีพจร"),
    col_("smoking", "สูบบุหรี่"),
    col_("alcohol", "แอลกอฮอล์"),
    col_("glucose", "น้ำตาล/FBS"),
    col_("cholesterol", "ไขมันรวม"),
    col_("phq2", "PHQ-2"),
    col_("salt", "ความเค็มในอาหาร"),
    col_("history", "ประวัติ/อาการสำคัญ"),
    col_("riskLevel", "ระดับความเสี่ยง"),
    col_("riskLabel", "กลุ่มผลแปลผล"),
    col_("flags", "ค่าผิดปกติ/ปัจจัยเสี่ยง"),
    col_("advice", "คำแนะนำรายบุคคล"),
    col_("followUp", "การติดตาม/ส่งต่อ"),
  ];
}

function getAnalysisColumns_() {
  return [
    col_("recordId", "รหัสรายการ"),
    col_("createdAt", "เวลาบันทึก"),
    col_("screeningDate", "วันที่คัดกรอง"),
    col_("monthKey", "เดือนรายงาน"),
    col_("fullName", "ชื่อ-สกุล"),
    col_("personId", "เลขบัตรประชาชน"),
    col_("age", "อายุ"),
    col_("gender", "เพศ"),
    col_("village", "หมู่บ้าน"),
    col_("houseNo", "บ้านเลขที่"),
    col_("volunteer", "อสม.ที่รับผิดชอบ"),
    col_("createdByName", "ผู้บันทึก"),
    col_("bmi", "BMI"),
    col_("bloodPressure", "ความดัน"),
    col_("glucose", "น้ำตาล/FBS"),
    col_("cholesterol", "ไขมันรวม"),
    col_("phq2", "PHQ-2"),
    col_("riskLevel", "ระดับความเสี่ยง"),
    col_("riskLabel", "กลุ่มผลแปลผล"),
    col_("abnormalFlags", "เหตุผล/ค่าผิดปกติ"),
    col_("aiAdvice", "ผลวิเคราะห์และคำแนะนำ"),
    col_("followUp", "แผนติดตาม"),
    col_("referralRequired", "ต้องส่งต่อ"),
  ];
}

function getReferralColumns_() {
  return [
    col_("recordId", "รหัสรายการ"),
    col_("createdAt", "เวลาบันทึก"),
    col_("screeningDate", "วันที่คัดกรอง"),
    col_("fullName", "ชื่อ-สกุล"),
    col_("personId", "เลขบัตรประชาชน"),
    col_("village", "หมู่บ้าน"),
    col_("houseNo", "บ้านเลขที่"),
    col_("volunteer", "อสม.ที่รับผิดชอบ"),
    col_("phone", "เบอร์โทร"),
    col_("riskLabel", "กลุ่มผลแปลผล"),
    col_("abnormalFlags", "เหตุผล/ค่าผิดปกติ"),
    col_("followUp", "คำแนะนำการส่งต่อ"),
    col_("referralStatus", "สถานะส่งต่อ/ติดตาม"),
    col_("referralDate", "วันที่ส่งต่อ/ติดตาม"),
    col_("serviceUnit", "หน่วยบริการ"),
    col_("note", "หมายเหตุ"),
  ];
}

function col_(key, label) {
  return { key: key, label: label };
}
