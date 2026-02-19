const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  // CHANNEL_PARTNER: "Channel Partener Lead FMS",
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
// FETCH LIST - Show when Planned (AM) is NOT NULL and Actual (AN) is NULL
// ============================================

async function getFilteredBookingLeads(sheets, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AQ`,  // Up to AQ (Unit No column)
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // Column mappings (0-based indices)
    // AM = column 39 (index 38) → Planned
    // AN = column 40 (index 39) → Actual
    // AO = column 41 (index 40) → Status
    // AP = column 42 (index 41) → Block
    // AQ = column 43 (index 42) → Unit No

    const plannedDate = row[38] ? row[38].trim() : "";         // AM → Planned
    const actualDate = row[39] ? row[39].trim() : "";          // AN → Actual
    const status = row[40] ? row[40].trim() : "";              // AO → Status
    const block = row[41] ? row[41].trim() : "";               // AP → Block
    const unitNo = row[42] ? row[42].trim() : "";              // AQ → Unit No

    // Show ONLY if:
    // - Planned date (AM) is NOT empty
    // - AND Actual date (AN) IS empty
    const showRow = plannedDate && !actualDate ;

    if (showRow) {
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
        block,
        unitNo,
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
      getFilteredBookingLeads(req.sheets, SHEETS.END_USER),
      getFilteredBookingLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    const all = [...endUser, ...channelPartner];

    // Sort by planned date
    all.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: all,
      total: all.length,
    });
  } catch (err) {
    console.error("Error fetching booking leads:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ============================================
// UPDATE - POST /update
// Handles: Status (AO), Block (AP), Unit No (AQ), Actual (AN)
// ============================================

router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, block, unitNo } = req.body;

    console.log("Booking Update Payload:", { sheetName, rowIndex, status, block, unitNo });

    if (!sheetName || !rowIndex) {
      return res.status(400).json({ success: false, error: "Missing sheetName or rowIndex" });
    }

    if (!status) {
      return res.status(400).json({ success: false, error: "Status is required" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Update Status (AO) - column 41
    updates.push({
      range: `'${sheetName}'!AO${rowIndex}`,
      values: [[String(status).trim()]],
    });

    // Update Block (AP) - column 42
    if (block && String(block).trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AP${rowIndex}`,
        values: [[String(block).trim()]],
      });
    }

    // Update Unit No (AQ) - column 43
    if (unitNo && String(unitNo).trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AQ${rowIndex}`,
        values: [[String(unitNo).trim()]],
      });
    }

    // Update Actual timestamp (AN) - marks booking as completed
    updates.push({
      range: `'${sheetName}'!AN${rowIndex}`,
      values: [[timestamp]],
    });

    // Batch update all changes
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    res.json({
      success: true,
      message: "Booking Done Successfully!",
    });
  } catch (err) {
    console.error("Booking update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;