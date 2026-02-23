const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEET_NAME = "END USER LEADS FMS";   // ← sirf yeh sheet ab kaam karegi

// ============================================
// Helpers (same as before)
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

  // agar already DD/MM/YYYY HH:mm:ss format me hai → use as is
  if (dateStr.includes(":") && dateStr.includes("/")) {
    return dateStr.trim();
  }

  // agar datetime-local (2026-02-26T14:44) aaya hai
  if (dateStr.includes("T")) {
    const [datePart, timePart] = dateStr.split("T");

    const [year, month, day] = datePart.split("-");
    const [hours, minutes] = timePart.split(":");

    return `${day}/${month}/${year} ${hours}:${minutes}:00`;
  }

  // fallback (sirf date ho)
  if (dateStr.includes("-")) {
    const [year, month, day] = dateStr.split("-");
    return `${day}/${month}/${year} 00:00:00`;
  }

  return dateStr;
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
// FETCH LIST - Only END USER sheet
// ============================================

async function getFilteredLeads(sheets) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A8:AL`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    const plannedDate = row[33] ? row[33].trim() : "";         // AH
    const actualDate = row[34] ? row[34].trim() : "";          // AI
    let status = row[35] ? row[35].trim() : "";                // AJ

    // All remarks columns
    const oldRemarks = row[11] ? row[11].trim() : "";
    const previousRemarksDate = row[13] ? row[13].trim() : "";
    const previousRemarks = row[19] ? row[19].trim() : "";
    const latestOldRemarksDate = row[21] ? row[21].trim() : "";
    const latestOldRemarks = row[24] ? row[24].trim() : "";
    const recentRemarksDate = row[27] ? row[27].trim() : "";
    const recentRemarks = row[32] ? row[32].trim() : "";
    const currentRemarks = row[37] ? row[37].trim() : "";

    let displayRemarks = currentRemarks || recentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

    const showRow = plannedDate &&
      (!status || status.trim().toLowerCase() === "rescheduled");

    if (showRow) {
      if (!status.trim()) status = "Pending";

      filtered.push({
        rowIndex: index + 8,
        sheetName: SHEET_NAME,          // fixed
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
    const leads = await getFilteredLeads(req.sheets);

    // Sort by planned date
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
    });
  } catch (err) {
    console.error("Error fetching leads:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

router.post("/update", async (req, res) => {
  try {
    const { rowIndex, status, rescheduleDate, remarks } = req.body;

    // sheetName ab frontend se nahi lenge — fixed hai
    const sheetName = SHEET_NAME;

    console.log("Update Payload:", { sheetName, rowIndex, status, rescheduleDate, remarks });

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
        range: `'${sheetName}'!Z${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn("Could not read followup count:", e.message);
    }
    const newFollowupCount = currentFollowupCount + 1;

    updates.push({
      range: `'${sheetName}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // Main update logic
    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      const newPlanned = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[newPlanned]],
      });
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } 
    else if (["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)) {
      updates.push({
        range: `'${sheetName}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,
        values: [[status]],
      });
      // Optional: clear planned date
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[""]],
      });
    } 
    else {
      updates.push({
        range: `'${sheetName}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // Save new remarks → AL
    if (remarks && String(remarks).trim() !== "") {
      updates.push({
        range: `'${sheetName}'!AL${rowIndex}`,
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
    console.error("Update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;