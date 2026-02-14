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
// FETCH LIST - Shows Pending + Rescheduled rows
// ============================================

async function getFilteredLeads(sheets, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AK`,  // extended to AK to be safe
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // 0-based indices
    const plannedDate = row[33] ? row[33].trim() : "";   // AG → index 32 (A=0, B=1, ..., AG=32)
    const actualDate = row[34] ? row[34].trim() : "";  // AH → index 33
    let status = row[35] ? row[35].trim() : "";  // AI → index 34
    const previousRemarks = row[32] ? row[32].trim() : ""; // AE or wherever previous remarks are

    // Show if planned exists AND it's not a final completed status
    const completedStatuses = ["Done", "Visit Done", "Completed", "Failed", "Cancelled", "Rejected"];
    const isCompleted = completedStatuses.some(s =>
      status.toLowerCase().includes(s.toLowerCase())
    );

    if (plannedDate && !isCompleted) {
      // Normalize empty status to Pending
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
        remarks: previousRemarks,
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

    if (!sheetName || !rowIndex) {
      return res.status(400).json({
        success: false,
        error: "sheetName and rowIndex are required",
      });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    if (rescheduleDate) {
      const newPlanned = getPlannedDateTime(rescheduleDate);

      // Planned → AG
      updates.push({
        range: `'${sheetName}'!AG${rowIndex}`,
        values: [[newPlanned]],
      });

      // Status → AI
      updates.push({
        range: `'${sheetName}'!AI${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } else {
      // Actual → AH
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[timestamp]],
      });

      // Status → AI
      updates.push({
        range: `'${sheetName}'!AI${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // Remarks → AK (only if provided and not empty)
    if (remarks !== undefined && remarks.trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AK${rowIndex}`,
        values: [[remarks.trim()]],
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
      message: rescheduleDate ? "Rescheduled Successfully" : "Updated Successfully",
    });
  } catch (err) {
    console.error("Update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;