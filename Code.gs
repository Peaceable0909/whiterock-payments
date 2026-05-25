// ================================================================
// WHITE ROCK — Counselor Payment Tracker  (Code.gs)
// ================================================================
// SETUP:
//  1. Extensions → Apps Script → paste this file
//  2. Run setupSheet() once to create headers + Counselors tab
//  3. Deploy → New Deployment → Web App
//     Execute as: Me  |  Who has access: Anyone
//  4. Copy Web App URL → paste into index.html SCRIPT_URL
// ================================================================

const SHEET_NAME     = "Payments";
const COUNSELOR_SHEET = "Counselors";

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmtDate(val) {
  if (!val || val === "") return "";
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch(e) { return String(val); }
}

// ── Get or create Payments sheet ───────────────────────────
function getPaymentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0 || sheet.getRange(1,1).getValue() !== "Counselor Name") {
    sheet.getRange(1,1,1,8).setValues([[
      "Counselor Name","Student Name","School / County",
      "Month","Payment Status","Amount","Date Paid","Submitted At"
    ]]);
    sheet.getRange(1,1,1,8).setFontWeight("bold").setBackground("#1e2f6e").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Get or create Counselors sheet ─────────────────────────
function getCounselorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COUNSELOR_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(COUNSELOR_SHEET);
    sheet.getRange(1,1).setValue("Counselor Name").setFontWeight("bold").setBackground("#1e2f6e").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Read all payment rows ───────────────────────────────────
function getAllRows() {
  try {
    const sheet = getPaymentsSheet();
    const last  = sheet.getLastRow();
    if (last < 2) return { rows: [] };
    const data = sheet.getRange(2, 1, last - 1, 8).getValues();
    const rows = data
      .filter(r => r[0] || r[1])
      .map(r => ({
        counselor   : String(r[0] || "").trim(),
        student     : String(r[1] || "").trim(),
        school      : String(r[2] || "").trim(),
        month       : String(r[3] || "").trim(),
        status      : String(r[4] || "").trim(),
        amount      : Number(r[5]) || 0,
        datePaid    : fmtDate(r[6]),
        submittedAt : fmtDate(r[7]),
      }));
    return { rows };
  } catch(err) { return { error: "getAllRows: " + err.toString() }; }
}

// ── Get counselor list ──────────────────────────────────────
function getCounselors() {
  try {
    const sheet = getCounselorsSheet();
    const last  = sheet.getLastRow();
    if (last < 2) return { counselors: [] };
    const data  = sheet.getRange(2, 1, last - 1, 1).getValues();
    const names = data.map(r => String(r[0] || "").trim()).filter(Boolean);
    return { counselors: names };
  } catch(err) { return { error: err.toString() }; }
}

// ── Add batch of payments ───────────────────────────────────
function addBatch(payments) {
  try {
    if (!Array.isArray(payments) || payments.length === 0)
      return { error: "No payments provided." };
    const sheet = getPaymentsSheet();
    const now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const rows  = payments.map(p => [
      String(p.counselor || "").trim(),
      String(p.student   || "").trim(),
      String(p.school    || "").trim(),
      String(p.month     || "").trim(),
      String(p.status    || "Paid").trim(),
      Number(p.amount)   || 0,
      p.datePaid ? String(p.datePaid) : today,
      now
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
    return { success: true, saved: rows.length, message: `${rows.length} payment(s) saved.` };
  } catch(err) { return { error: "addBatch: " + err.toString() }; }
}

function addPayment(p) { return addBatch([p]); }

// ── Summary for admin ───────────────────────────────────────
function getSummary() {
  const { rows, error } = getAllRows();
  if (error) return { error };
  const MONTHS   = ["April 2026", "May 2026"];
  const names    = [...new Set(rows.map(r => r.counselor).filter(Boolean))].sort();
  const schools  = [...new Set(rows.map(r => r.school).filter(Boolean))].sort();
  const stats    = names.map(name => {
    const mine = rows.filter(r => r.counselor === name);
    return {
      name,
      totalPaid    : mine.filter(r => r.status === "Paid").length,
      totalPending : mine.filter(r => r.status === "Pending").length,
      totalAmount  : mine.filter(r => r.status === "Paid").reduce((s,r) => s+r.amount, 0),
      byMonth      : MONTHS.map(m => {
        const mr = mine.filter(r => r.month === m && r.status === "Paid");
        return { month:m, count:mr.length, amount:mr.reduce((s,r)=>s+r.amount,0), schools:[...new Set(mr.map(r=>r.school))] };
      })
    };
  });
  const paid = rows.filter(r => r.status === "Paid");
  return {
    counselorStats  : stats,
    grandTotal      : { count:paid.length, amount:paid.reduce((s,r)=>s+r.amount,0) },
    schoolBreakdown : schools.map(sc => {
      const sr = paid.filter(r => r.school === sc);
      return { school:sc, count:sr.length, amount:sr.reduce((s,r)=>s+r.amount,0) };
    }).sort((a,b) => b.count - a.count),
    months      : MONTHS,
    lastUpdated : new Date().toISOString()
  };
}

// ── doGet ───────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "getSummary";
    if (action === "getSummary")   return jsonOut(getSummary());
    if (action === "getAllData")   return jsonOut(getAllRows());
    if (action === "getCounselors") return jsonOut(getCounselors());
    if (action === "ping")         return jsonOut({ ok:true, time:new Date().toISOString() });
    return jsonOut({ error: "Unknown action: " + action });
  } catch(err) { return jsonOut({ error: "doGet: " + err.toString() }); }
}

// ── doPost ──────────────────────────────────────────────────
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents)
      return jsonOut({ error: "No POST body." });
    const body   = JSON.parse(e.postData.contents);
    if (Array.isArray(body))          return jsonOut(addBatch(body));
    const action = body.action || "addPayment";
    if (action === "addBatch")        return jsonOut(addBatch(body.payments || []));
    if (action === "addPayment")      return jsonOut(addPayment(body));
    return jsonOut({ error: "Unknown action: " + action });
  } catch(err) { return jsonOut({ error: "doPost: " + err.toString() }); }
}

// ── SETUP ───────────────────────────────────────────────────
// Run this ONCE manually in the Apps Script editor.
function setupSheet() {
  getPaymentsSheet();
  const cs = getCounselorsSheet();
  // Add counselors if sheet is empty
  if (cs.getLastRow() < 2) {
    cs.getRange(2,1,15,1).setValues([
      ["Mrs Lydia"],
      ["Michael"],
      ["Zainab"],
      ["Peaceable"],
      ["Trust"],
      ["Mr Anya"],
      ["Jude Scott"],
      ["Praise"],
      ["Mrs Gladys"],
      ["Feranmi"],
      ["Mr Adegoke"],
      ["Robert"],
      ["Mr Abraham"],
      ["Cherish"],
      ["Mrs Oluwo"],["Jason"],["Daniel"]
    ]);
  }
  SpreadsheetApp.getUi().alert("✅ Setup complete! Edit the Counselors tab to add/remove counselors.");
}

function setupSampleData() {
  const sample = [
    ["Mrs Lydia",   "Sample Student 1","Lincoln High / Cook County",  "April 2026","Paid",   150,"2026-04-03"],
    ["Michael",     "Sample Student 2","Riverside HS / Lake County",  "April 2026","Paid",   200,"2026-04-05"],
    ["Zainab",      "Sample Student 3","Northside HS / Cook County",  "May 2026",  "Paid",   175,"2026-05-07"],
    ["Peaceable",   "Sample Student 4","Westfield MS / Will County",  "May 2026",  "Pending",160,""],
    ["Trust",       "Sample Student 5","Jefferson MS / DuPage County","April 2026","Paid",   120,"2026-04-14"],
  ];
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const rows = sample.map(r => [...r, now]);
  const sheet = getPaymentsSheet();
  sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 8).setValues(rows);
  SpreadsheetApp.getUi().alert("✅ Sample data added.");
}
