const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// ============================================
// Helpers
// ============================================

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

// Helper: Reschedule ke liye Date+Time merge karna
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const timePart = `${hours}:${minutes}:${seconds}`;

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-"); 
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  return `${formattedDate} ${timePart}`;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// ============================================
// FETCH LIST (GET)
// ============================================

async function getFilteredLeads(sheets, sheetName) {
  // AJ tak data fetch karenge (AJ = Index 35)
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AJ`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // --- COLUMN MAPPING (Based on Screenshot Step 5) ---
    // AE = 30 (Previous/Followup Remarks) -> READ THIS
    // AF = 31 (Planned)
    // AG = 32 (Actual)
    // AH = 33 (Status)
    // AJ = 35 (Current Meeting Remarks) -> WRITE HERE

    const plannedDate = row[31] ? row[31].trim() : "";   // Column AF
    const actualDate = row[32] ? row[32].trim() : "";    // Column AG
    const status = row[33] ? row[33].trim() : "";        // Column AH
    
    // CHANGE: Hamein AE (30) padhna hai list dikhane ke liye
    const previousRemarks = row[30] ? row[30].trim() : ""; 

    // Filter: Show only if Planned exists AND Actual is empty
    if (plannedDate && !actualDate) {
      filtered.push({
        rowIndex: index + 8,
        sheetName,
        uniqueId: row[1] || "",
        customerName: row[2] || "",
        customerContact: row[3] || "",
        interestedIn: row[4] || "",
        projectSelection: row[5] || "",
        plannedDate,
        status: status || "Pending",
        remarks: previousRemarks, // Frontend ko Column AE bhej rahe hain
      });
    }
  });

  return filtered;
}

// ============================================
// ROUTES
// ============================================

router.get("/list", async (req, res) => {
  try {
    const [endUser, channelPartner] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    const all = [...endUser, ...channelPartner];

    all.sort((a, b) => {
      const da = parseDate(a.plannedDate);
      const db = parseDate(b.plannedDate);
      return da - db;
    });

    res.json({
      success: true,
      data: all,
      total: all.length,
    });
  } catch (err) {
    console.error("Error fetching Step 5 data:", err);
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

// ============================================
// UPDATE (POST)
// ============================================

router.post("/update", async (req, res) => {
  try {
    const {
      sheetName,
      rowIndex,
      status,
      rescheduleDate,
      remarks,
    } = req.body;

    if (!sheetName || !rowIndex) {
      return res.status(400).json({
        success: false,
        error: "sheetName and rowIndex are required",
      });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // --- SCENARIO 1: RESCHEDULE ---
    if (rescheduleDate) {
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      
      // Update Planned (Col AF)
      updates.push({
        range: `'${sheetName}'!AF${rowIndex}`,
        values: [[newPlannedDateTime]],
      });
      
      // Status Update (Optional, keeping consistent)
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [["Rescheduled"]], 
      });

    } 
    // --- SCENARIO 2: DONE / FAILED ---
    else {
      // Update Actual (Col AG)
      updates.push({
        range: `'${sheetName}'!AG${rowIndex}`,
        values: [[timestamp]],
      });

      // Update Status (Col AH)
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[status]],
      });
    }

    // --- COMMON: UPDATE REMARKS (Col AJ) ---
    // Naya remark hamesha Column AJ mein jayega
    if (remarks !== undefined) {
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,
        values: [[remarks]],
      });
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    res.json({
      success: true,
      message: rescheduleDate ? "Rescheduled Successfully" : "Status Updated Successfully",
    });

  } catch (err) {
    console.error("❌ Error updating Step 5:", err);
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

module.exports = router;