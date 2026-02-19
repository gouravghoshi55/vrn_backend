const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = "END USER LEADS FMS";

// ============================================
// Helpers (copied & slightly cleaned)
// ============================================

function getCurrentTimestamp() {
  const now = new Date();
  return [
    String(now.getDate()).padStart(2, "0"),
    String(now.getMonth() + 1).padStart(2, "0"),
    now.getFullYear(),
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ].join("/").replace(/(\d{2}\/\d{2}\/\d{4})\/(\d{2}:\d{2}:\d{2})/, "$1 $2");
}

function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }

  const now = new Date();
  const timePart = [
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ].join(":");

  return `${formattedDate} ${timePart}`;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// ============================================
// FETCH Filtered Leads
// ============================================

async function getFilteredNBDLeads(sheets) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!A8:AL`,
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, index) => {
      const plannedDate    = row[33]?.trim() || "";   // AH (34)
      const actualDate     = row[34]?.trim() || "";   // AI (35)
      let   status         = row[35]?.trim() || "";   // AJ (36)

      // All remarks columns
      const oldRemarks          = row[11]?.trim() || "";     // L
      const previousRemarksDate = row[13]?.trim() || "";     // N
      const previousRemarks     = row[19]?.trim() || "";     // T
      const latestOldRemarksDate= row[21]?.trim() || "";     // V
      const latestOldRemarks    = row[24]?.trim() || "";     // Y
      const recentRemarksDate   = row[27]?.trim() || "";     // AB
      const recentRemarks       = row[32]?.trim() || "";     // AG
      const currentRemarks      = row[37]?.trim() || "";     // AL

      // Display logic: newest → oldest
      const displayRemarks = currentRemarks || recentRemarks || latestOldRemarks ||
                             previousRemarks || oldRemarks;

      const showRow = !!plannedDate &&
        (!status || status.trim().toLowerCase() === "rescheduled");

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
  } catch (err) {
    console.error(`Error fetching END USER leads:`, err.message);
    throw err;
  }
}

// ============================================
// ROUTES
// ============================================

router.get("/nbd/list", async (req, res) => {
  try {
    console.log("Fetching END USER follow-up leads...");
    const leads = await getFilteredNBDLeads(req.sheets);

    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "end-user",
    });
  } catch (err) {
    console.error("Error fetching END USER leads:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

router.post("/nbd/update", async (req, res) => {
  try {
    const { rowIndex, status, rescheduleDate, remarks } = req.body;

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    console.log("Updating END USER lead:", { rowIndex, status, rescheduleDate });

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Get & increment Followup Count (column Z)
    let currentCount = 0;
    try {
      const res = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${SHEET_NAME}'!Z${rowIndex}`,
      });
      currentCount = parseInt(res.data.values?.[0]?.[0] || "0", 10);
    } catch {}

    const newCount = currentCount + 1;
    updates.push({
      range: `'${SHEET_NAME}'!Z${rowIndex}`,
      values: [[newCount]],
    });

    if (rescheduleDate && rescheduleDate.trim()) {
      // Reschedule
      const newPlanned = getPlannedDateTime(rescheduleDate);
      updates.push({ range: `'${SHEET_NAME}'!AH${rowIndex}`, values: [[newPlanned]] });
      updates.push({ range: `'${SHEET_NAME}'!AJ${rowIndex}`, values: [["Rescheduled"]] });
    } else if (["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)) {
      // Final close cases
      updates.push({ range: `'${SHEET_NAME}'!AI${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${SHEET_NAME}'!AJ${rowIndex}`, values: [[status]] });
      // Optional: clear planned date
      updates.push({ range: `'${SHEET_NAME}'!AH${rowIndex}`, values: [[""]] });
    } else {
      // Mark done / other
      updates.push({ range: `'${SHEET_NAME}'!AI${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${SHEET_NAME}'!AJ${rowIndex}`, values: [[status || "Done"]] });
    }

    if (remarks?.trim()) {
      updates.push({
        range: `'${SHEET_NAME}'!AL${rowIndex}`,
        values: [[remarks.trim()]],
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
      message: rescheduleDate ? "Rescheduled" : 
               ["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status) 
                 ? `Marked as ${status}` : "Updated",
      newFollowupCount: newCount,
    });
  } catch (err) {
    console.error("END USER update failed:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;