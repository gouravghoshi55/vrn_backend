const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

/*
W (22) = Planned
X (23) = Actual
Y (24) = Status
Z (25) = Deal Meeting Date
AA (26) = Next FollowUp Date
AB (27) = FollowUp Count
AC (28) = Remarks
*/

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
// FETCH LIST
// ============================================

async function getFilteredLeads(sheets, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AC`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    const plannedDate = row[22] ? row[22].trim() : "";
    const actualDate = row[23] ? row[23].trim() : "";
    const status = row[24] ? row[24].trim() : "";
    const followUpCount = row[27] ? row[27].trim() : "0";

    const statusLower = status.toLowerCase();

    // ✅ Planned exists AND
    // Actual empty OR status is no conversation
    if (plannedDate && (!actualDate || statusLower === "no conversation")) {
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
        followUpCount: parseInt(followUpCount) || 0,
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
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

// ============================================
// UPDATE
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
    // NO CONVERSATION
    // ----------------------------
    if (normalizedStatus === "no conversation") {
      const newCount = (parseInt(currentFollowUpCount) || 0) + 1;

      if (nextFollowUpDate) {
        updates.push({
          range: `'${sheetName}'!W${rowIndex}`,
          values: [[nextFollowUpDate]],
        });
      }

      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[timestamp]],
      });

      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [["No Conversation"]],
      });

      if (nextFollowUpDate) {
        updates.push({
          range: `'${sheetName}'!AA${rowIndex}`,
          values: [[nextFollowUpDate]],
        });
      }

      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[newCount.toString()]],
      });

      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[remarks]],
        });
      }
    }

    // ----------------------------
    // DONE
    // ----------------------------
    else if (normalizedStatus === "done") {
      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[timestamp]],
      });

      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [["Done"]],
      });

      if (dealMeetingDate) {
        updates.push({
          range: `'${sheetName}'!Z${rowIndex}`,
          values: [[dealMeetingDate]],
        });
      }

      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[remarks]],
        });
      }
    }

    // ----------------------------
    // NOT INTERESTED
    // ----------------------------
    else if (normalizedStatus === "not interested") {
      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[timestamp]],
      });

      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [["Not Interested"]],
      });

      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
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
      message: "After Field Visit Updated",
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message,
    });
  }
});

module.exports = router;
