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

// Helper: Merges User Selected Date with Current Time for 'Planned' column
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
  // Fetch up to Column AF (Remarks)
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AF`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // COLUMN MAPPING BASED ON STEP 4 SCREENSHOT
    // Z=25 (Planned), AA=26 (Actual), AB=27 (Status)
    // AE=30 (FollowUp Count), AF=31 (Remarks)
    
    const plannedDate = row[25] ? row[25].trim() : "";   // Column Z
    const actualDate = row[26] ? row[26].trim() : "";    // Column AA
    const status = row[27] ? row[27].trim() : "";        // Column AB
    const followUpCount = row[30] ? row[30].trim() : "0";// Column AE
    const remarks = row[31] ? row[31].trim() : "";       // Column AF

    // ============================================================
    // UPDATED CONDITION:
    // Show row ONLY IF: Planned is NOT Empty AND Actual IS Empty
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
        plannedDate,
        status: status || "Pending",
        followUpCount: parseInt(followUpCount) || 0,
        remarks: remarks,
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
    console.error("Error fetching Step 4 data:", err);
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
      dealMeetingDate,
      nextFollowUpDate,
      remarks,
      currentFollowUpCount,
    } = req.body;

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "sheetName, rowIndex, status required",
      });
    }

    const normalizedStatus = status.trim().toLowerCase();
    const timestamp = getCurrentTimestamp();
    const updates = [];

    // ----------------------------
    // CASE 1: NO CONVERSATION (Loop Back)
    // ----------------------------
    if (normalizedStatus === "no conversation") {
      const newCount = (parseInt(currentFollowUpCount) || 0) + 1;
      
      // 1. Update Planned (Col Z) -> New Date
      if (nextFollowUpDate) {
        const plannedDateTime = getPlannedDateTime(nextFollowUpDate);
        updates.push({
          range: `'${sheetName}'!Z${rowIndex}`,
          values: [[plannedDateTime]],
        });
      }

      // 2. Update Actual (Col AA) -> Timestamp
      updates.push({
        range: `'${sheetName}'!AA${rowIndex}`,
        values: [[timestamp]],
      });

      // 3. Update Status (Col AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [["No conversation"]],
      });

      // 4. Update Next FollowUp Date (Col AD)
      if (nextFollowUpDate) {
        updates.push({
          range: `'${sheetName}'!AD${rowIndex}`,
          values: [[nextFollowUpDate]],
        });
      }

      // 5. Update FollowUp Count (Col AE)
      updates.push({
        range: `'${sheetName}'!AE${rowIndex}`,
        values: [[newCount.toString()]],
      });

      // 6. Update Remarks (Col AF)
      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AF${rowIndex}`,
          values: [[remarks]],
        });
      }
    }

    // ----------------------------
    // CASE 2: DONE
    // ----------------------------
    else if (normalizedStatus === "done") {
      // 1. Update Actual (Col AA)
      updates.push({
        range: `'${sheetName}'!AA${rowIndex}`,
        values: [[timestamp]],
      });

      // 2. Update Status (Col AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [["Done"]],
      });

      // 3. Update Deal Meeting Date (Col AC)
      if (dealMeetingDate) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[dealMeetingDate]],
        });
      }

      // 4. Update Remarks (Col AF)
      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AF${rowIndex}`,
          values: [[remarks]],
        });
      }
    }

    // ----------------------------
    // CASE 3: OTHERS (Not Interested etc.)
    // ----------------------------
    else {
      // 1. Update Actual (Col AA)
      updates.push({
        range: `'${sheetName}'!AA${rowIndex}`,
        values: [[timestamp]],
      });

      // 2. Update Status (Col AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[status]],
      });

      // 3. Update Remarks (Col AF)
      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AF${rowIndex}`,
          values: [[remarks]],
        });
      }
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
      message: "Call/Deal Follow Up Updated Successfully",
    });
  } catch (err) {
    console.error("❌ Error updating Step 4:", err);
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

module.exports = router;