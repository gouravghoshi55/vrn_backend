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

// Helper: Merges User Selected Date with Current Time for 'Planned' column (Reschedule ke liye)
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const timePart = `${hours}:${minutes}:${seconds}`;

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-"); // Expecting YYYY-MM-DD
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
  // Fetch up to Column AJ (Remarks)
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AJ`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // UPDATED INDICES BASED ON SCREENSHOT (Step 5)
    // AF=31 (Planned), AG=32 (Actual), AH=33 (Status), AJ=35 (Remarks)
    
    const plannedDate = row[31] ? row[31].trim() : "";   // Column AF
    const actualDate = row[32] ? row[32].trim() : "";    // Column AG
    const status = row[33] ? row[33].trim() : "";        // Column AH
    const remarks = row[35] ? row[35].trim() : "";       // Column AJ

    // ============================================================
    // CONDITION: Show row ONLY IF: Planned is NOT Empty AND Actual IS Empty
    // ============================================================
    if (plannedDate && !actualDate) {
      filtered.push({
        rowIndex: index + 8,
        sheetName,
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
        remarks: remarks,
      });
    }
  });

  return filtered;
}

// ============================================
// ROUTES
// ============================================

/**
 * GET /api/meeting-nbd/list
 */
router.get("/list", async (req, res) => {
  try {
    const [endUser, channelPartner] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    const all = [...endUser, ...channelPartner];

    // Sort by Planned Date
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
    console.error("Error fetching Meeting NBD data:", err);
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

// ============================================
// UPDATE (POST)
// ============================================

/**
 * POST /api/meeting-nbd/update
 * 
 * Supports:
 * 1. Reschedule: Updates Planned (AF) with new Date + Time.
 * 2. Done/Status Update: Updates Actual (AG) with Timestamp and Status (AH).
 */
router.post("/update", async (req, res) => {
  try {
    const {
      sheetName,
      rowIndex,
      status,
      rescheduleDate, // Only required if Rescheduling
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

    // --- MAPPING (Step 5) ---
    // AF (31) = Planned
    // AG (32) = Actual
    // AH (33) = Status
    // AJ (35) = Remarks

    // ----------------------------
    // SCENARIO 1: RESCHEDULE
    // ----------------------------
    if (rescheduleDate) {
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      
      // Update Planned (Col AF) - Overwrite with new date
      updates.push({
        range: `'${sheetName}'!AF${rowIndex}`,
        values: [[newPlannedDateTime]],
      });

      // (Optional) Update Status to 'Reschedule' or keep it pending
      // Hum Status (AH) bhi update kar sakte hain taaki pata chale
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [["Rescheduled"]], 
      });

    } 
    // ----------------------------
    // SCENARIO 2: DONE / NEGOTIATION FAILED / ETC
    // ----------------------------
    else {
      // 1. Update Actual (Col AG) -> Timestamp
      updates.push({
        range: `'${sheetName}'!AG${rowIndex}`,
        values: [[timestamp]],
      });

      // 2. Update Status (Col AH)
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[status]],
      });
    }

    // ----------------------------
    // COMMON: UPDATE REMARKS (Col AJ)
    // ----------------------------
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

    console.log(`✅ Meeting NBD Updated: Row ${rowIndex}, Rescheduled: ${!!rescheduleDate}`);

    res.json({
      success: true,
      message: rescheduleDate ? "Meeting Rescheduled Successfully" : "Meeting Status Updated Successfully",
    });

  } catch (err) {
    console.error("❌ Error updating Meeting NBD:", err);
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

module.exports = router;