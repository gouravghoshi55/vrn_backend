const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CP_SHEET_NAME = "Channel Partener Lead FMS"; // ✅ Only CP sheet

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
// READ DATA - CP After Field Visit with Category Filter
// ======================================================

async function getFilteredCPLeads(sheets, category) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${CP_SHEET_NAME}'!A8:AR`, // Extended to AR
    });

    const rows = response.data.values || [];

    return rows
      .map((row, index) => {
        const getCol = (idx) => (row[idx] ? String(row[idx]).trim() : "");

        const plannedDate = getCol(26); // AA
        const actualDate = getCol(27);  // AB
        const status = getCol(28);      // AC
        const canContact = getCol(43);  // AR

        // Category filter
        const categoryMatch =
          (category === "can-contact" && canContact === "Yes") ||
          (category === "cannot-contact" && canContact === "No");

        if (!categoryMatch) return null;

        // Filter condition
        if ((plannedDate && !actualDate) || status === "No conversation" || status === "Next Follow Up") {
          // ===== ALL REMARKS COLUMNS =====
          const oldRemarks = getCol(11);           // L
          const previousRemarksDate = getCol(13);  // N
          const previousRemarks = getCol(19);      // T
          const latestOldRemarks = getCol(24);     // Y
          const latestOldRemarksDate = getCol(21); // V
          const currentRemarks = getCol(32);       // AG

          let displayRemarks = currentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

          return {
            rowIndex: index + 8,
            sheetName: CP_SHEET_NAME,
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
            canContact: canContact,
          };
        }
        return null;
      })
      .filter(Boolean);
  } catch (error) {
    console.error(`Error fetching CP After Field Visit (${category}):`, error.message);
    throw error;
  }
}

// ======================================================
// GET ENDPOINTS - CP After Field Visit
// ======================================================

// Can Contact List
router.get("/can-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Can Contact After Field Visit data...");

    const leads = await getFilteredCPLeads(req.sheets, "can-contact");

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
      category: "can-contact",
    });
  } catch (error) {
    console.error("❌ Error:", error);
    res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

// Cannot Contact List
router.get("/cannot-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Cannot Contact After Field Visit data...");

    const leads = await getFilteredCPLeads(req.sheets, "cannot-contact");

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
      category: "cannot-contact",
    });
  } catch (error) {
    console.error("❌ Error:", error);
    res.status(500).json({
      success: false,
      message: error.message,
    });
  }
});

// ======================================================
// POST ENDPOINTS - CP After Field Visit Update
// ======================================================

// Can Contact Update
router.post("/can-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate, dealMeetingDate } = req.body;

    console.log("📝 CP Can Contact After Field Visit Update:", {
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
        range: `'${CP_SHEET_NAME}'!AF${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentCount = parseInt(val) || 0;
    } catch (e) {
      console.warn("Could not read count:", e.message);
    }

    const newCount = currentCount + 1;
    const updates = [];
    const timestamp = getCurrentTimestamp();

    updates.push({
      range: `'${CP_SHEET_NAME}'!AF${rowIndex}`,
      values: [[newCount]],
    });

    if (remarks && String(remarks).trim()) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AG${rowIndex}`,
        values: [[String(remarks).trim()]],
      });
    }

    if (
      rescheduleDate &&
      String(rescheduleDate).trim() !== "" &&
      ["No conversation", "Next Follow Up"].includes(status || "")
    ) {
      const formatted = formatDateToSheetStyle(rescheduleDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!AA${rowIndex}`,
        values: [[formatted]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AE${rowIndex}`,
        values: [[formatted]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AC${rowIndex}`,
        values: [[status]],
      });
    } else {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AB${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AC${rowIndex}`,
        values: [[status || "Done"]],
      });
      if (dealMeetingDate) {
        const formattedDeal = formatDateToSheetStyle(dealMeetingDate);
        updates.push({
          range: `'${CP_SHEET_NAME}'!AD${rowIndex}`,
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
    }

    res.json({
      success: true,
      message: "Updated successfully",
      newFollowUpCount: newCount,
    });
  } catch (error) {
    console.error("❌ Update failed:", error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// Cannot Contact Update
router.post("/cannot-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate, dealMeetingDate } = req.body;

    console.log("📝 CP Cannot Contact After Field Visit Update:", {
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
        range: `'${CP_SHEET_NAME}'!AF${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentCount = parseInt(val) || 0;
    } catch (e) {
      console.warn("Could not read count:", e.message);
    }

    const newCount = currentCount + 1;
    const updates = [];
    const timestamp = getCurrentTimestamp();

    updates.push({
      range: `'${CP_SHEET_NAME}'!AF${rowIndex}`,
      values: [[newCount]],
    });

    if (remarks && String(remarks).trim()) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AG${rowIndex}`,
        values: [[String(remarks).trim()]],
      });
    }

    if (
      rescheduleDate &&
      String(rescheduleDate).trim() !== "" &&
      ["No conversation", "Next Follow Up"].includes(status || "")
    ) {
      const formatted = formatDateToSheetStyle(rescheduleDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!AA${rowIndex}`,
        values: [[formatted]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AE${rowIndex}`,
        values: [[formatted]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AC${rowIndex}`,
        values: [[status]],
      });
    } else {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AB${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AC${rowIndex}`,
        values: [[status || "Done"]],
      });
      if (dealMeetingDate) {
        const formattedDeal = formatDateToSheetStyle(dealMeetingDate);
        updates.push({
          range: `'${CP_SHEET_NAME}'!AD${rowIndex}`,
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
    }

    res.json({
      success: true,
      message: "Updated successfully",
      newFollowUpCount: newCount,
    });
  } catch (error) {
    console.error("❌ Update failed:", error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

module.exports = router;