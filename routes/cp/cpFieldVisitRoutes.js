const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CP_SHEET_NAME = "Channel Partner Lead FMS"; // ✅ Only CP sheet

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
// READ DATA - CP Field Visit with Category Filter
// ======================================================

async function getFilteredCPLeads(sheets, category) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${CP_SHEET_NAME}'!A8:AT`, // Extended to AT for Can Contact column
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      const importantNote = row[10] ? row[10].trim() : "";     // K
      const plannedDate = row[20] ? row[20].trim() : "";       // U
      const actualDate = row[21] ? row[21].trim() : "";        // V
      const status = row[22] ? row[22].trim() : "";            // W

      // ===== REMARKS COLUMNS =====
      const oldRemarks = row[11] ? row[11].trim() : "";        // L
      const previousRemarksDate = row[13] ? row[13].trim() : ""; // N
      const previousRemarks = row[19] ? row[19].trim() : "";   // T
      const latestRemarks = row[24] ? row[24].trim() : "";     // Y
      const followupCount = row[25] ? parseInt(row[25].trim(), 10) || 0 : 0; // Z

      const canContact = row[45] ? row[45].trim() : ""; // Column AT (Can Contact)

      // ===== CATEGORY FILTER =====
      const categoryMatch =
        (category === "can-contact" && canContact === "Yes") ||
        (category === "cannot-contact" && (canContact === "No" || canContact ===""));
      if (!categoryMatch) return;

      // Display logic: Y > T > L
      let displayRemarks = "";
      if (latestRemarks) {
        displayRemarks = latestRemarks;
      } else if (previousRemarks) {
        displayRemarks = previousRemarks;
      } else if (oldRemarks) {
        displayRemarks = oldRemarks;
      }

      // Show only if status is empty or "Rescheduled"
      const showRow = !status || status.trim().toLowerCase() === "rescheduled";

      if (showRow && plannedDate) {
        filteredLeads.push({
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
          importantNote: importantNote,
          plannedDate: plannedDate,
          status: status || "Pending",
          followupCount: followupCount,
          remarks: displayRemarks,
          oldRemarks: oldRemarks,
          previousRemarks: previousRemarks,
          previousRemarksDate: previousRemarksDate,
          canContact: canContact,
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching CP Field Visit (${category}):`, error.message);
    throw error;
  }
}

// ======================================================
// GET ENDPOINTS - CP Field Visit
// ======================================================

// Can Contact List
router.get("/can-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Can Contact Field Visit data...");

    const leads = await getFilteredCPLeads(req.sheets, "can-contact");

    leads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
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
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// Cannot Contact List
router.get("/cannot-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Cannot Contact Field Visit data...");

    const leads = await getFilteredCPLeads(req.sheets, "cannot-contact");

    leads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
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
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// ======================================================
// POST ENDPOINTS - CP Field Visit Update
// ======================================================

// Can Contact Update
router.post("/can-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate } = req.body;

    console.log("📝 CP Can Contact Field Visit Update:", { rowIndex, status, rescheduleDate, remarks });

    if (!rowIndex) {
      return res.status(400).json({
        success: false,
        error: "Missing rowIndex",
      });
    }

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // Fetch current Followup Count
    let currentFollowupCount = 0;
    try {
      const countResponse = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      });
      const val = countResponse.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn(`⚠️ Could not read followup count:`, e.message);
    }

    const newFollowupCount = currentFollowupCount + 1;

    // Always increment Followup Count
    updates.push({
      range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // Main update logic
    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!U${rowIndex}`,
        values: [[newPlannedDateTime]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!W${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } else if (status === "Not Interested") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!W${rowIndex}`,
        values: [["Not Interested"]],
      });
    } else {
      updates.push({
        range: `'${CP_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!W${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // Remarks
    if (remarks && String(remarks).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!Y${rowIndex}`,
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

    console.log(`✅ CP Can Contact Field Visit updated: Row ${rowIndex}`);

    res.json({
      success: true,
      message: rescheduleDate
        ? "Visit Rescheduled successfully"
        : status === "Not Interested"
          ? "Marked as Not Interested"
          : "Field Visit marked as Done",
      newFollowupCount: newFollowupCount,
    });
  } catch (error) {
    console.error("❌ Update failed:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

// Cannot Contact Update
router.post("/cannot-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate } = req.body;

    console.log("📝 CP Cannot Contact Field Visit Update:", { rowIndex, status, rescheduleDate, remarks });

    if (!rowIndex) {
      return res.status(400).json({
        success: false,
        error: "Missing rowIndex",
      });
    }

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // Fetch current Followup Count
    let currentFollowupCount = 0;
    try {
      const countResponse = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      });
      const val = countResponse.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn(`⚠️ Could not read followup count:`, e.message);
    }

    const newFollowupCount = currentFollowupCount + 1;

    updates.push({
      range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!U${rowIndex}`,
        values: [[newPlannedDateTime]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!W${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } else if (status === "Not Interested") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!W${rowIndex}`,
        values: [["Not Interested"]],
      });
    } else {
      updates.push({
        range: `'${CP_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!W${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    if (remarks && String(remarks).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!Y${rowIndex}`,
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

    console.log(`✅ CP Cannot Contact Field Visit updated: Row ${rowIndex}`);

    res.json({
      success: true,
      message: rescheduleDate
        ? "Visit Rescheduled successfully"
        : status === "Not Interested"
          ? "Marked as Not Interested"
          : "Field Visit marked as Done",
      newFollowupCount: newFollowupCount,
    });
  } catch (error) {
    console.error("❌ Update failed:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

module.exports = router;