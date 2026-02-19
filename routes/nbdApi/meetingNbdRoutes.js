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
// FETCH LIST - Only show when Status is empty or "Rescheduled"
// ============================================

async function getFilteredLeads(sheets, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AL`,  // Extended to AL (column 38)
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // Column mappings (0-based indices)
    const plannedDate = row[33] ? row[33].trim() : "";         // AH → Planned (column 34)
    const actualDate = row[34] ? row[34].trim() : "";          // AI → Actual (column 35)
    let status = row[35] ? row[35].trim() : "";                // AJ → Status (column 36)
    
    // ===== ALL REMARKS COLUMNS =====
    const oldRemarks = row[11] ? row[11].trim() : "";          // L - Initial/Oldest Remarks
    const previousRemarksDate = row[13] ? row[13].trim() : ""; // N - Date for T remarks
    const previousRemarks = row[19] ? row[19].trim() : "";     // T - Previous Remarks
    const latestOldRemarksDate = row[21] ? row[21].trim() : "";// V - Date for Y remarks
    const latestOldRemarks = row[24] ? row[24].trim() : "";    // Y - Latest Old Remarks
    const recentRemarksDate = row[27] ? row[27].trim() : "";   // AB - Date for AG remarks
    const recentRemarks = row[32] ? row[32].trim() : "";       // AG - Recent Remarks
    const currentRemarks = row[37] ? row[37].trim() : "";      // AL - Current/New Remarks
    
    // Display priority for table: AL > AG > Y > T > L
    let displayRemarks = currentRemarks || recentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

    // Show ONLY if:
    // - Planned date exists
    // - AND status is empty OR exactly "Rescheduled" (case-insensitive)
    const showRow = plannedDate &&
      (!status || status.trim().toLowerCase() === "rescheduled");

    if (showRow) {
      // Normalize empty status to "Pending"
      if (!status.trim()) status = "Pending";

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
        status,
        
        // ===== SEND ALL REMARKS TO FRONTEND =====
        remarks: displayRemarks,                      // For table display
        oldRemarks: oldRemarks,                       // L - Initial remarks (read-only)
        previousRemarks: previousRemarks,             // T - Previous remarks (read-only)
        previousRemarksDate: previousRemarksDate,     // N - Date for T
        latestOldRemarks: latestOldRemarks,          // Y - Latest old remarks (read-only)
        latestOldRemarksDate: latestOldRemarksDate,  // V - Date for Y
        recentRemarks: recentRemarks,                // AG - Recent remarks (read-only)
        recentRemarksDate: recentRemarksDate,        // AB - Date for AG
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
    const [endUser, channelPartner] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    const all = [...endUser, ...channelPartner];

    all.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: all,
      total: all.length,
    });
  } catch (err) {
    console.error("Error fetching leads:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ============================================
// UPDATE - POST /update
// ============================================

router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, rescheduleDate, remarks } = req.body;

    console.log("Update Payload:", { sheetName, rowIndex, status, rescheduleDate, remarks });

    if (!sheetName || !rowIndex) {
      return res.status(400).json({ success: false, error: "Missing sheetName or rowIndex" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Fetch current Followup Count from Z (har update par +1)
    let currentFollowupCount = 0;
    try {
      const countRes = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!Z${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn("Could not read followup count:", e.message);
    }
    const newFollowupCount = currentFollowupCount + 1;

    // Always increment Followup Count (Z)
    updates.push({
      range: `'${sheetName}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // Main logic
    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      // RESCHEDULE
      const newPlanned = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,  // Planned
        values: [[newPlanned]],
      });
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,  // Status
        values: [["Rescheduled"]],
      });
    } 
    else if (["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)) {
      // NEW: Negotiation Failed / Deal Not Done / Not Interested
      console.log(`→ Processing FINAL CLOSE: ${status}`);

      // Actual timestamp (AI)
      updates.push({
        range: `'${sheetName}'!AI${rowIndex}`,
        values: [[timestamp]],
      });

      // Status (AJ)
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,
        values: [[status]],
      });

      // Optional: Planned date clear kar do taaki list se hat jaye
      // Agar chahte ho to yeh line comment kar dena
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[""]],
      });
    } 
    else {
      // MARK DONE or other
      updates.push({
        range: `'${sheetName}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // ===== NEW REMARKS SAVE TO COLUMN AL =====
    if (remarks && String(remarks).trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AL${rowIndex}`,  // New remarks save to AL
        values: [[String(remarks).trim()]],
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
      message: rescheduleDate 
        ? "Rescheduled Successfully" 
        : ["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)
          ? "Marked as " + status
          : "Updated Successfully",
      newFollowupCount: newFollowupCount,
    });
  } catch (err) {
    console.error("Update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;