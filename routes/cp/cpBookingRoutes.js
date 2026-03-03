const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CP_SHEET_NAME = "Channel Partner Lead FMS"; // ✅ Only CP sheet

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
// FETCH LIST - CP Booking with Category Filter
// ============================================

async function getFilteredCPBookingLeads(sheets, category) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${CP_SHEET_NAME}'!A8:AT`, // Extended to AT for Can Contact column
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // 0-based indices
    const plannedDate = row[38] ? row[38].trim() : "";   // AM → Planned
    const actualDate = row[39] ? row[39].trim() : "";    // AN → Actual
    const status = row[40] ? row[40].trim() : "";        // AO → Status
    const block = row[41] ? row[41].trim() : "";         // AP → Block
    const unitNo = row[42] ? row[42].trim() : "";        // AQ → Unit No
    const canContact = row[45] ? row[45].trim() : "";    // AT → Can Contact

    // ===== CATEGORY FILTER =====
    const categoryMatch =
      (category === "can-contact" && canContact === "Yes") ||
      (category === "cannot-contact" && canContact === "No");

    if (!categoryMatch) return;

    // Show row only if Planned date exists
    const showRow = plannedDate !== "";

    if (showRow) {
      filtered.push({
        rowIndex: index + 8,
        sheetName: CP_SHEET_NAME,
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
        canContact,
      });
    }
  });

  return filtered;
}

// ============================================
// GET ENDPOINTS
// ============================================

// Can Contact Booking List
router.get("/can-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Can Contact Booking data...");

    const leads = await getFilteredCPBookingLeads(req.sheets, "can-contact");

    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "can-contact",
    });
  } catch (err) {
    console.error("❌ Error fetching CP Can Contact Booking:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Cannot Contact Booking List
router.get("/cannot-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Cannot Contact Booking data...");

    const leads = await getFilteredCPBookingLeads(req.sheets, "cannot-contact");

    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "cannot-contact",
    });
  } catch (err) {
    console.error("❌ Error fetching CP Cannot Contact Booking:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ============================================
// POST ENDPOINTS - CP Booking Update
// ============================================

// Can Contact Update
router.post("/can-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, block, unitNo } = req.body;

    console.log("📝 CP Can Contact Booking Update:", { rowIndex, status, block, unitNo });

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    if (!status) {
      return res.status(400).json({ success: false, error: "Status is required" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Status (AO)
    updates.push({
      range: `'${CP_SHEET_NAME}'!AO${rowIndex}`,
      values: [[String(status).trim()]],
    });

    // Block (AP)
    if (block && String(block).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AP${rowIndex}`,
        values: [[String(block).trim()]],
      });
    }

    // Unit No (AQ)
    if (unitNo && String(unitNo).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AQ${rowIndex}`,
        values: [[String(unitNo).trim()]],
      });
    }

    // Actual timestamp (AN)
    updates.push({
      range: `'${CP_SHEET_NAME}'!AN${rowIndex}`,
      values: [[timestamp]],
    });

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ CP Can Contact Booking updated: Row ${rowIndex}`);

    res.json({
      success: true,
      message: "Booking Done Successfully!",
    });
  } catch (err) {
    console.error("❌ CP Can Contact Booking update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Cannot Contact Update
router.post("/cannot-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, block, unitNo } = req.body;

    console.log("📝 CP Cannot Contact Booking Update:", { rowIndex, status, block, unitNo });

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    if (!status) {
      return res.status(400).json({ success: false, error: "Status is required" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Status (AO)
    updates.push({
      range: `'${CP_SHEET_NAME}'!AO${rowIndex}`,
      values: [[String(status).trim()]],
    });

    // Block (AP)
    if (block && String(block).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AP${rowIndex}`,
        values: [[String(block).trim()]],
      });
    }

    // Unit No (AQ)
    if (unitNo && String(unitNo).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AQ${rowIndex}`,
        values: [[String(unitNo).trim()]],
      });
    }

    // Actual timestamp (AN)
    updates.push({
      range: `'${CP_SHEET_NAME}'!AN${rowIndex}`,
      values: [[timestamp]],
    });

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ CP Cannot Contact Booking updated: Row ${rowIndex}`);

    res.json({
      success: true,
      message: "Booking Done Successfully!",
    });
  } catch (err) {
    console.error("❌ CP Cannot Contact Booking update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;