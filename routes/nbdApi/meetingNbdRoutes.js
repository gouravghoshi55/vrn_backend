const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = "END USER LEADS FMS";

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

  if (dateStr.includes(":") && dateStr.includes("/")) {
    return dateStr.trim();
  }

  if (dateStr.includes("T")) {
    const [datePart, timePart] = dateStr.split("T");
    const [year, month, day] = datePart.split("-");
    const [hours, minutes] = timePart.split(":");
    return `${day}/${month}/${year} ${hours}:${minutes}:00`;
  }

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
// FETCH LIST
// ============================================

async function getFilteredLeads(sheets) {
  // ✅ CHANGED: Extended to AU (46)
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A8:AU`, // ← Changed from AL to AU
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    const plannedDate = row[33] ? row[33].trim() : "";
    const actualDate = row[34] ? row[34].trim() : "";
    let status = row[35] ? row[35].trim() : "";

    const oldRemarks = row[11] ? row[11].trim() : "";
    const previousRemarksDate = row[13] ? row[13].trim() : "";
    const previousRemarks = row[19] ? row[19].trim() : "";
    const latestOldRemarksDate = row[21] ? row[21].trim() : "";
    const latestOldRemarks = row[24] ? row[24].trim() : "";
    const recentRemarksDate = row[27] ? row[27].trim() : "";
    const recentRemarks = row[32] ? row[32].trim() : "";
    const currentRemarks = row[37] ? row[37].trim() : "";

    let displayRemarks =
      currentRemarks || recentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

    const showRow =
      plannedDate &&
      (!status ||
        status.trim().toLowerCase() === "rescheduled" ||
        status.trim().toLowerCase() === "next field visit required");

    if (showRow) {
      if (!status.trim()) status = "Pending";

      filtered.push({
        rowIndex: index + 8,
        sheetName: SHEET_NAME,
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
        fieldVisitCount: row[46] || "0", // ✅ NEW: AU column (index 46)
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
    const { rowIndex, status, rescheduleDate, nextFieldVisitDate, remarks } = req.body;

    const sheetName = SHEET_NAME;

    console.log("Update Payload:", {
      sheetName,
      rowIndex,
      status,
      rescheduleDate,
      nextFieldVisitDate,
      remarks,
    });

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // ✅ Followup Count (Z column) - unchanged
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

    // ✅ NEW: Field Visit Count (AU column)
    if (status === "Next Field Visit Required") {
      let currentFieldVisitCount = 0;
      try {
        const fieldVisitRes = await req.sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${sheetName}'!AU${rowIndex}`,
        });
        const val = fieldVisitRes.data.values?.[0]?.[0];
        currentFieldVisitCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
      } catch (e) {
        console.warn("Could not read field visit count:", e.message);
      }

      const newFieldVisitCount = currentFieldVisitCount + 1;

      updates.push({
        range: `'${sheetName}'!AU${rowIndex}`,
        values: [[newFieldVisitCount]],
      });

      console.log(
        `✅ Field Visit Count incremented: ${currentFieldVisitCount} → ${newFieldVisitCount}`
      );
    }

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
    // ✅ NEW: Handle Next Field Visit Required
    else if (status === "Next Field Visit Required" && nextFieldVisitDate && String(nextFieldVisitDate).trim() !== "") {
      const newPlanned = getPlannedDateTime(nextFieldVisitDate);
      
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`, // Update planned date
        values: [[newPlanned]],
      });
      
      updates.push({
        range: `'${sheetName}'!AJ${rowIndex}`, // Update status
        values: [["Next Field Visit Required"]],
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
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[""]],
      });
    } else {
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
        : status === "Next Field Visit Required"
        ? "Next Field Visit Scheduled Successfully"
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