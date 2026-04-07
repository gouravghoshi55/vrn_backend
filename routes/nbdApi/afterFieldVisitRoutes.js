const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

function getCurrentTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2,"0"), mo = String(now.getMonth()+1).padStart(2,"0"), y = now.getFullYear();
  const h = String(now.getHours()).padStart(2,"0"), mi = String(now.getMinutes()).padStart(2,"0"), s = String(now.getSeconds()).padStart(2,"0");
  return `${d}/${mo}/${y} ${h}:${mi}:${s}`;
}

function getShortTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2,"0"), mo = String(now.getMonth()+1).padStart(2,"0"), y = now.getFullYear();
  const h = String(now.getHours()).padStart(2,"0"), mi = String(now.getMinutes()).padStart(2,"0");
  return `[${d}/${mo}/${y} ${h}:${mi}]`;
}

function formatDateToSheetStyle(dateInput) {
  if (!dateInput) return ""; const dt = new Date(dateInput); if (isNaN(dt.getTime())) return "";
  const d = String(dt.getDate()).padStart(2,"0"), m = String(dt.getMonth()+1).padStart(2,"0"), y = dt.getFullYear();
  const h = String(dt.getHours()).padStart(2,"0"), mi = String(dt.getMinutes()).padStart(2,"0"), s = String(dt.getSeconds()).padStart(2,"0");
  return `${d}/${m}/${y} ${h}:${mi}:${s}`;
}

function getDoerTag(user) { if (!user) return null; if (user.role==="admin"||user.assignedModule==="all") return null; if (user.assignedModule==="fsr") return null; const m={"bdm1@company.com":"BDM1","bdm2@company.com":"BDM2","bdm3@company.com":"BDM3"}; return m[user.email?.toLowerCase()]||null; }
function getFSRDoerTag(user) { if (!user) return null; if (user.assignedModule!=="fsr") return null; const m={"bdm4@company.com":"BDM4","bdm5@company.com":"BDM5"}; return m[user.email?.toLowerCase()]||null; }

async function buildAppendedRemarks(sheets, sheetName, cellRange, newRemark) {
  if (!newRemark || !String(newRemark).trim()) return null;
  let existing = "";
  try { const r = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${sheetName}'!${cellRange}` }); existing = r.data.values?.[0]?.[0] || ""; } catch (e) {}
  const timestamped = `${getShortTimestamp()} ${String(newRemark).trim()}`;
  return existing.trim() ? `${existing.trim()}\n${timestamped}` : timestamped;
}

async function appendToLogger(sheets, leadInfo, stepName, status, remarks, userEmail) {
  try { const ts = getCurrentTimestamp(); await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: `'${LOGGER_SHEET_NAME}'!A:J`, valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS", requestBody: { values: [[ts, stepName, leadInfo.uniqueId||"", leadInfo.customerName||"", leadInfo.customerContact||"", leadInfo.interestedIn||"", leadInfo.projectSelection||"", status, remarks||"", userEmail||""]] } }); } catch (e) { console.error("❌ Logger failed:", e.message); }
}

async function appendToNotInterestedSheet(sheets, leadInfo, stepName, reason, userEmail) {
  try { const ts = getCurrentTimestamp(); await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: `'${NOT_INTERESTED_SHEET}'!A:K`, valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS", requestBody: { values: [[ts, stepName, leadInfo.uniqueId||"", leadInfo.customerName||"", leadInfo.customerContact||"", leadInfo.interestedIn||"", leadInfo.projectSelection||"", leadInfo.leadSource||"", leadInfo.doer||"", reason||"", userEmail||""]] } }); } catch (e) { console.error("❌ NI sheet failed:", e.message); }
}

async function getFilteredLeads(sheets, user) {
  try {
    const response = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${NBD_SHEET_NAME}'!A8:AN` });
    const rows = response.data.values || [];
    const doerTag = getDoerTag(user), fsrDoerTag = getFSRDoerTag(user);

    return rows.map((row, index) => {
      const g = (i) => (row[i] ? String(row[i]).trim() : "");
      const plannedDate = g(26), actualDate = g(27), status = g(28), doer = g(38), fsrDoer = g(39);
      const showRow = (plannedDate && !actualDate) || status === "No conversation" || status === "Next Follow Up" || status === "Next Field Visit Required";
      if (!showRow) return null;
      if (doerTag && doer !== doerTag) return null;
      if (fsrDoerTag && fsrDoer !== fsrDoerTag) return null;

      const oldRemarks = g(11), previousRemarksDate = g(13), previousRemarks = g(19);
      const latestOldRemarks = g(24), latestOldRemarksDate = g(21), currentRemarks = g(32);

      return {
        rowIndex: index+8, sheetName: NBD_SHEET_NAME,
        uniqueId: g(1), customerName: g(2), customerContact: g(3),
        interestedIn: g(4), projectSelection: g(5), leadSource: g(6),
        leadGenNumber: g(7), leadGenName: g(8), importantNote: g(10),
        plannedDate, status: status || "Pending", followUpCount: g(31)||"0",
        fsrDoer, remarks: currentRemarks || latestOldRemarks || previousRemarks || oldRemarks,
        oldRemarks, previousRemarks, previousRemarksDate, latestOldRemarks, latestOldRemarksDate, doer,
      };
    }).filter(Boolean);
  } catch (error) { throw error; }
}

router.get("/list", async (req, res) => {
  try {
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => { const p = (d) => { if (!d) return 0; const pp = d.split(/[\/ :]/); return new Date(pp[2], pp[1]-1, pp[0], pp[3]||0, pp[4]||0).getTime(); }; return p(a.plannedDate) - p(b.plannedDate); });
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) { res.status(500).json({ success: false, message: error.message }); }
});

router.post("/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate, dealMeetingDate, leadInfo, notInterestedReason } = req.body;
    if (!rowIndex) return res.status(400).json({ success: false, error: "Missing rowIndex" });

    let currentCount = 0;
    try { const cr = await req.sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${NBD_SHEET_NAME}'!AF${rowIndex}` }); currentCount = parseInt(cr.data.values?.[0]?.[0]) || 0; } catch (e) {}
    const newCount = currentCount + 1;
    const updates = []; const timestamp = getCurrentTimestamp();

    updates.push({ range: `'${NBD_SHEET_NAME}'!AF${rowIndex}`, values: [[newCount]] });

    // ✅ Remarks — APPEND with timestamp
    if (remarks && String(remarks).trim()) {
      const appended = await buildAppendedRemarks(req.sheets, NBD_SHEET_NAME, `AG${rowIndex}`, remarks);
      if (appended) updates.push({ range: `'${NBD_SHEET_NAME}'!AG${rowIndex}`, values: [[appended]] });
    }

    if (status === "Not Interested") {
      updates.push({ range: `'${NBD_SHEET_NAME}'!AB${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, values: [["Not Interested"]] });
      if (leadInfo) await appendToNotInterestedSheet(req.sheets, leadInfo, "Step 3 - After Field Visit", notInterestedReason||"", req.user?.email);
    } else if (rescheduleDate && String(rescheduleDate).trim() !== "" && ["No conversation", "Next Follow Up", "Next Field Visit Required"].includes(status||"")) {
      const formatted = formatDateToSheetStyle(rescheduleDate);
      updates.push({ range: `'${NBD_SHEET_NAME}'!AA${rowIndex}`, values: [[formatted]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AE${rowIndex}`, values: [[formatted]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, values: [[status]] });
    } else {
      updates.push({ range: `'${NBD_SHEET_NAME}'!AB${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, values: [[status || "Done"]] });
      if (dealMeetingDate) updates.push({ range: `'${NBD_SHEET_NAME}'!AD${rowIndex}`, values: [[formatDateToSheetStyle(dealMeetingDate)]] });
    }

    await req.sheets.spreadsheets.values.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { valueInputOption: "USER_ENTERED", data: updates } });
    if (leadInfo) await appendToLogger(req.sheets, leadInfo, "Step 3 - After Field Visit", status||"Done", remarks||"", req.user?.email);

    res.json({ success: true, message: "Updated successfully", newFollowUpCount: newCount });
  } catch (error) { res.status(500).json({ success: false, error: error.message }); }
});

module.exports = router;