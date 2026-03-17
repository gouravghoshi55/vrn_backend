const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = "END USER LEADS FMS";

function getCurrentTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0");
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const s = String(now.getSeconds()).padStart(2, "0");
  return `${d}/${m}/${y} ${h}:${mi}:${s}`;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// ✅ Helper: Get doerTag from user info
function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  const emailToDoerMap = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
  };
  return emailToDoerMap[user.email?.toLowerCase()] || null;
}

// ============================================
// FETCH LIST
// ============================================

async function getFilteredBookingLeads(sheets, user) {
  // ✅ Extended range from A8:AQ to A8:AS to include Doer column
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A8:AS`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  // ✅ Get doer tag for current user
  const doerTag = getDoerTag(user);

  rows.forEach((row, index) => {
    const plannedDate = row[38] ? row[38].trim() : "";   // AM
    const actualDate  = row[39] ? row[39].trim() : "";   // AN
    const status      = row[40] ? row[40].trim() : "";   // AO
    const block       = row[41] ? row[41].trim() : "";   // AP
    const unitNo      = row[42] ? row[42].trim() : "";   // AQ
    const doer        = row[45] ? row[45].trim() : "";   // ✅ AS = index 44

    const showRow = plannedDate !== "";

    if (showRow) {
      // ✅ DOER FILTER
      if (doerTag && doer !== doerTag) return;

      filtered.push({
        rowIndex: index + 8,
        sheetName: SHEET_NAME,
        uniqueId: row[1] || "",
        customerName: row[2] || "",
        customerContact: row[3] || "",
        interestedIn: row[4] || "",
        projectSelection: row[5] || "",
        leadSource: row[6] || "",
        leadGenNumber: row[7] || "",
        leadGenName: row[8] || "",
        plannedDate,
        status: status || "Pending",
        block,
        unitNo,
        doer, // ✅ Include doer
      });
    }
  });

  return filtered;
}

// ============================================
// ROUTES - GET /list
// ============================================

router.get("/list", async (req, res) => {
  try {
    // ✅ Pass user info
    const leads = await getFilteredBookingLeads(req.sheets, req.user);
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (err) {
    console.error("Error fetching booking leads:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ============================================
// UPDATE - POST /update
// ============================================

router.post("/update", async (req, res) => {
  try {
    const { rowIndex, status, block, unitNo } = req.body;
    const sheetName = SHEET_NAME;

    if (!rowIndex) return res.status(400).json({ success: false, error: "Missing rowIndex" });
    if (!status) return res.status(400).json({ success: false, error: "Status is required" });

    const timestamp = getCurrentTimestamp();
    const updates = [];

    updates.push({ range: `'${sheetName}'!AO${rowIndex}`, values: [[String(status).trim()]] });

    if (block && String(block).trim() !== "") {
      updates.push({ range: `'${sheetName}'!AP${rowIndex}`, values: [[String(block).trim()]] });
    }

    if (unitNo && String(unitNo).trim() !== "") {
      updates.push({ range: `'${sheetName}'!AQ${rowIndex}`, values: [[String(unitNo).trim()]] });
    }

    updates.push({ range: `'${sheetName}'!AN${rowIndex}`, values: [[timestamp]] });

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    res.json({ success: true, message: "Booking Done Successfully!" });
  } catch (err) {
    console.error("Booking update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// NOTE: Assign endpoint is centralized in nbdinRoutes.js (/leads/nbdin/assign)

module.exports = router;