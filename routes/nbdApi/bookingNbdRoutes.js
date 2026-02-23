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
    if (parts[0].length === 4) {
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// ============================================
// FETCH LIST - Only Planned (AM) is NOT NULL
// ============================================

async function getFilteredBookingLeads(sheets) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A8:AQ`,  // Up to AQ (Unit No)
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // 0-based indices
    const plannedDate  = row[38] ? row[38].trim() : "";   // AM → Planned (col 39)
    const actualDate   = row[39] ? row[39].trim() : "";   // AN → Actual (col 40)
    const status       = row[40] ? row[40].trim() : "";   // AO → Status (col 41)
    const block        = row[41] ? row[41].trim() : "";   // AP → Block (col 42)
    const unitNo       = row[42] ? row[42].trim() : "";   // AQ → Unit No (col 43)

    // Show row only if Planned date exists (AM not empty)
    // You can add more strict condition if needed: && !actualDate
    const showRow = plannedDate !== "";

    if (showRow) {
      filtered.push({
        rowIndex: index + 8,
        sheetName: SHEET_NAME,           // fixed
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
    const leads = await getFilteredBookingLeads(req.sheets);

    // Sort by planned date
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
    });
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

    // sheetName ab fixed hai — frontend se nahi bhejna padega
    const sheetName = SHEET_NAME;

    console.log("Booking Update Payload:", { sheetName, rowIndex, status, block, unitNo });

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    if (!status) {
      return res.status(400).json({ success: false, error: "Status is required" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Status (AO) – always update
    updates.push({
      range: `'${sheetName}'!AO${rowIndex}`,
      values: [[String(status).trim()]],
    });

    // Block (AP) – optional
    if (block && String(block).trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AP${rowIndex}`,
        values: [[String(block).trim()]],
      });
    }

    // Unit No (AQ) – optional
    if (unitNo && String(unitNo).trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AQ${rowIndex}`,
        values: [[String(unitNo).trim()]],
      });
    }

    // Actual timestamp (AN) – mark as done
    updates.push({
      range: `'${sheetName}'!AN${rowIndex}`,
      values: [[timestamp]],
    });

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