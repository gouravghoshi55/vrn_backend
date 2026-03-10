const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS"; // ✅ Only END USER sheet

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

function formatDateToSheetStyle(dateInput) {
  if (!dateInput) return "";
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) return "";
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");
  const seconds = String(d.getSeconds()).padStart(2, "0");
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

// ======================================================
// READ DATA - NBD After Field Visit
// ======================================================

async function getFilteredLeads(sheets) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AG`,
    });

    const rows = response.data.values || [];

    return rows
      .map((row, index) => {
        const getCol = (idx) => (row[idx] ? String(row[idx]).trim() : "");

        const plannedDate = getCol(26); // AA
        const actualDate = getCol(27);  // AB
        const status = getCol(28);      // AC

        // Filter: show if planned exists and actual empty OR specific status
        if ((plannedDate && !actualDate) || status === "No conversation" || status === "Next Follow Up" || status ==="Next Field Visit Required") {
          // ===== ALL REMARKS COLUMNS =====
          const oldRemarks = getCol(11);           // L
          const previousRemarksDate = getCol(13);  // N
          const previousRemarks = getCol(19);      // T
          const latestOldRemarks = getCol(24);     // Y
          const latestOldRemarksDate = getCol(21); // V
          const currentRemarks = getCol(32);       // AG

          // Display priority: AG > Y > T > L
          let displayRemarks = currentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

          return {
            rowIndex: index + 8,
            sheetName: NBD_SHEET_NAME,
            uniqueId: getCol(1),
            customerName: getCol(2),
            customerContact: getCol(3),
            interestedIn: getCol(4),
            projectSelection: getCol(5),
            leadSource: getCol(6),
            leadGenNumber: getCol(7),
            leadGenName: getCol(8),
            importantNote: getCol(10),
            plannedDate,
            status: status || "Pending",
            followUpCount: getCol(31) || "0", // AF
            remarks: displayRemarks,
            oldRemarks: oldRemarks,
            previousRemarks: previousRemarks,
            previousRemarksDate: previousRemarksDate,
            latestOldRemarks: latestOldRemarks,
            latestOldRemarksDate: latestOldRemarksDate,
          };
        }
        return null;
      })
      .filter(Boolean);
  } catch (error) {
    console.error("Error fetching NBD After Field Visit leads:", error.message);
    throw error;
  }
}

// ======================================================
// GET /list - NBD After Field Visit
// ======================================================

router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching NBD After Field Visit data (END USER only)...");

    const leads = await getFilteredLeads(req.sheets);

    leads.sort((a, b) => {
      const parse = (d) => {
        if (!d) return 0;
        const p = d.split(/[\/ :]/);
        return new Date(p[2], p[1] - 1, p[0], p[3] || 0, p[4] || 0).getTime();
      };
      return parse(a.plannedDate) - parse(b.plannedDate);
    });

    res.json({
      success: true,
      data: leads,
      total: leads.length,
    });
  } catch (error) {
    console.error("❌ Error fetching NBD After Field Visit:", error);
    res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

// ======================================================
// POST /update - NBD After Field Visit
// ======================================================

router.post("/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate, dealMeetingDate } = req.body;

    console.log("📝 NBD After Field Visit Update:", {
      rowIndex,
      status,
      remarks,
      rescheduleDate,
      dealMeetingDate,
    });

    if (!rowIndex) {
      return res.status(400).json({
        success: false,
        error: "Missing rowIndex",
      });
    }

    let currentCount = 0;
    try {
      const countRes = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${NBD_SHEET_NAME}'!AF${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentCount = parseInt(val) || 0;
    } catch (e) {
      console.warn("Could not read count:", e.message);
    }

    const newCount = currentCount + 1;
    const updates = [];
    const timestamp = getCurrentTimestamp();

    // Always increment count (AF)
    updates.push({
      range: `'${NBD_SHEET_NAME}'!AF${rowIndex}`,
      values: [[newCount]],
    });

    // Always update new remarks (AG)
    if (remarks && String(remarks).trim()) {
      updates.push({
        range: `'${NBD_SHEET_NAME}'!AG${rowIndex}`,
        values: [[String(remarks).trim()]],
      });
    }

    // RESCHEDULE CASE
    if (
      rescheduleDate &&
      String(rescheduleDate).trim() !== "" &&
      ["No conversation", "Next Follow Up"].includes(status || "")
    ) {
      const formatted = formatDateToSheetStyle(rescheduleDate);
      console.log("→ Processing RESCHEDULE to:", formatted);

      updates.push({
        range: `'${NBD_SHEET_NAME}'!AA${rowIndex}`, // Planned
        values: [[formatted]],
      });

      updates.push({
        range: `'${NBD_SHEET_NAME}'!AE${rowIndex}`, // Next FollowUp Date
        values: [[formatted]],
      });

      updates.push({
        range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, // Status
        values: [[status]],
      });
    }
    // MARK DONE / OTHER STATUS
    else {
      console.log("→ Processing DONE / STATUS UPDATE:", status);

      // Actual timestamp (AB)
      updates.push({
        range: `'${NBD_SHEET_NAME}'!AB${rowIndex}`,
        values: [[timestamp]],
      });

      // Status (AC)
      let finalStatus = status || "Done";
      updates.push({
        range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`,
        values: [[finalStatus]],
      });

      // Deal Meeting Date (AD)
      if (dealMeetingDate) {
        const formattedDeal = formatDateToSheetStyle(dealMeetingDate);
        updates.push({
          range: `'${NBD_SHEET_NAME}'!AD${rowIndex}`,
          values: [[formattedDeal]],
        });
      }
    }

    if (updates.length > 0) {
      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: updates,
        },
      });
      console.log(`✅ NBD After Field Visit updated: Row ${rowIndex}`);
    }

    res.json({
      success: true,
      message: "Updated successfully",
      newFollowUpCount: newCount,
    });
  } catch (error) {
    console.error("❌ NBD After Field Visit update failed:", error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

module.exports = router;