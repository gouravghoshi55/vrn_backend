const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// ======================================================
// HELPERS
// ======================================================

function getCurrentTimestamp() {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const year = now.getFullYear();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

// Handles "YYYY-MM-DD" and "YYYY-MM-DDTHH:mm"
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";

  if (dateStr.includes("T")) {
    const [datePart, timePart] = dateStr.split("T");
    const [year, month, day] = datePart.split("-");
    return `${day}/${month}/${year} ${timePart}:00`;
  }

  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }

  return `${formattedDate} ${hours}:${minutes}:${seconds}`;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[/-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// ======================================================
// READ DATA
// ======================================================

async function getFilteredLeads(sheets, sheetName) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:X`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      const importantNote = row[10] ? row[10].trim() : "";

      const plannedDate = row[20] ? row[20].trim() : "";
      const actualDate = row[21] ? row[21].trim() : "";
      const status = row[22] ? row[22].trim() : "";
      // =============================
      // REMARKS LOGIC (T first, then Y)
      // =============================

      // T = index 19
      // Y = index 24
      const remarkFromT = row[19] ? row[19].trim() : "";
      const remarkFromY = row[24] ? row[24].trim() : "";

      let finalRemarks = "";

      // First priority -> T
      if (remarkFromT) {
        finalRemarks = remarkFromT;
      }
      // If T empty -> Y
      else if (remarkFromY) {
        finalRemarks = remarkFromY;
      }
      // Filter: Planned exists AND Actual empty
      if (plannedDate && !actualDate) {
        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: sheetName,
          uniqueId: row[1] || "",
          customerName: row[2] || "",
          customerContact: row[3] || "",
          interestedIn: row[4] || "",
          projectSelection: row[5] || "",
          leadSource: row[6] || "",
          leadGenNumber: row[7] || "",
          leadGenName: row[8] || "",
          importantNote: importantNote,
          plannedDate: plannedDate,
          status: status || "Pending",
          remarks: finalRemarks,
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
}

// ======================================================
// ROUTES
// ======================================================

router.get("/list", async (req, res) => {
  try {
    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    allLeads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });

    res.json({ success: true, data: allLeads, total: allLeads.length });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// ======================================================
// UPDATE
// ======================================================

router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks, rescheduleDate } = req.body;

    console.log("📝 Updating:", { sheetName, rowIndex, status, rescheduleDate });

    if (!sheetName || !rowIndex) {
      return res
        .status(400)
        .json({ success: false, error: "Missing required fields" });
    }

    const updates = [];

    // RESCHEDULE
    if (rescheduleDate) {
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${sheetName}'!T${rowIndex}`,
        values: [[newPlannedDateTime]],
      });
    }
    // MARK DONE
    else {
      const timestamp = getCurrentTimestamp();
      updates.push({
        range: `'${sheetName}'!U${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${sheetName}'!V${rowIndex}`,
        values: [[status]],
      });
    }

    // UPDATE REMARKS -> COLUMN X
    if (remarks !== undefined) {
      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
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
      message: rescheduleDate
        ? "Visit Rescheduled successfully"
        : "Field Visit marked as Done",
    });
  } catch (error) {
    console.error("❌ Error updating:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

module.exports = router;