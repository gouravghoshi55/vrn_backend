const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CP_SHEET_NAME = "Channel Partener Lead FMS"; // ✅ Only CP sheet

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
// FETCH LIST - CP Meeting with Category Filter
// ============================================

async function getFilteredCPLeads(sheets, category) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${CP_SHEET_NAME}'!A8:AT`, // Extended to AT for Can Contact column
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    const plannedDate = row[33] ? row[33].trim() : "";         // AH
    const actualDate = row[34] ? row[34].trim() : "";          // AI
    let status = row[35] ? row[35].trim() : "";                // AJ
    const canContact = row[45] ? row[45].trim() : "";          // AT

    // ===== CATEGORY FILTER =====
    const categoryMatch =
      (category === "can-contact" && canContact === "Yes") ||
      (category === "cannot-contact" && canContact === "No");

    if (!categoryMatch) return;

    // All remarks columns
    const oldRemarks = row[11] ? row[11].trim() : "";          // L
    const previousRemarksDate = row[13] ? row[13].trim() : ""; // N
    const previousRemarks = row[19] ? row[19].trim() : "";     // T
    const latestOldRemarksDate = row[21] ? row[21].trim() : "";// V
    const latestOldRemarks = row[24] ? row[24].trim() : "";    // Y
    const recentRemarksDate = row[27] ? row[27].trim() : "";   // AB
    const recentRemarks = row[32] ? row[32].trim() : "";       // AG
    const currentRemarks = row[37] ? row[37].trim() : "";      // AL

    let displayRemarks = currentRemarks || recentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

    const showRow = plannedDate &&
      (!status || status.trim().toLowerCase() === "rescheduled");

    if (showRow) {
      if (!status.trim()) status = "Pending";

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
        status,
        remarks: displayRemarks,
        oldRemarks,
        previousRemarks,
        previousRemarksDate,
        latestOldRemarks,
        latestOldRemarksDate,
        recentRemarks,
        recentRemarksDate,
        canContact,
      });
    }
  });

  return filtered;
}

// ============================================
// GET ENDPOINTS
// ============================================

// Can Contact List
router.get("/can-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Can Contact Meeting data...");

    const leads = await getFilteredCPLeads(req.sheets, "can-contact");

    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "can-contact",
    });
  } catch (err) {
    console.error("❌ Error fetching CP Can Contact Meeting:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Cannot Contact List
router.get("/cannot-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Cannot Contact Meeting data...");

    const leads = await getFilteredCPLeads(req.sheets, "cannot-contact");

    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "cannot-contact",
    });
  } catch (err) {
    console.error("❌ Error fetching CP Cannot Contact Meeting:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ============================================
// POST ENDPOINTS - CP Meeting Update
// ============================================

// Can Contact Update
router.post("/can-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, rescheduleDate, remarks } = req.body;

    console.log("📝 CP Can Contact Meeting Update:", { rowIndex, status, rescheduleDate, remarks });

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Followup Count (Z column) increment
    let currentFollowupCount = 0;
    try {
      const countRes = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn("Could not read followup count:", e.message);
    }
    const newFollowupCount = currentFollowupCount + 1;

    updates.push({
      range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // Main update logic
    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      const newPlanned = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!AH${rowIndex}`,
        values: [[newPlanned]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AJ${rowIndex}`,
        values: [["Rescheduled"]],
      });
    }
    else if (["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AJ${rowIndex}`,
        values: [[status]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AH${rowIndex}`,
        values: [[""]],
      });
    }
    else {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AJ${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // Save new remarks → AL
    if (remarks && String(remarks).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AL${rowIndex}`,
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

    res.json({
      success: true,
      message: rescheduleDate
        ? "Rescheduled Successfully"
        : ["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)
          ? "Marked as " + status
          : "Updated Successfully",
      newFollowupCount,
    });
  } catch (err) {
    console.error("❌ CP Can Contact Meeting update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Cannot Contact Update
router.post("/cannot-contact/update", async (req, res) => {
  try {
    const { rowIndex, status, rescheduleDate, remarks } = req.body;

    console.log("📝 CP Cannot Contact Meeting Update:", { rowIndex, status, rescheduleDate, remarks });

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Followup Count (Z column) increment
    let currentFollowupCount = 0;
    try {
      const countRes = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn("Could not read followup count:", e.message);
    }
    const newFollowupCount = currentFollowupCount + 1;

    updates.push({
      range: `'${CP_SHEET_NAME}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // Main update logic
    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      const newPlanned = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!AH${rowIndex}`,
        values: [[newPlanned]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AJ${rowIndex}`,
        values: [["Rescheduled"]],
      });
    }
    else if (["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AJ${rowIndex}`,
        values: [[status]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AH${rowIndex}`,
        values: [[""]],
      });
    }
    else {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${CP_SHEET_NAME}'!AJ${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // Save new remarks → AL
    if (remarks && String(remarks).trim() !== "") {
      updates.push({
        range: `'${CP_SHEET_NAME}'!AL${rowIndex}`,
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

    res.json({
      success: true,
      message: rescheduleDate
        ? "Rescheduled Successfully"
        : ["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)
          ? "Marked as " + status
          : "Updated Successfully",
      newFollowupCount,
    });
  } catch (err) {
    console.error("❌ CP Cannot Contact Meeting update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;