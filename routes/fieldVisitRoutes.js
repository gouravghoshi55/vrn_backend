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
// READ DATA (Filter: show only when status is empty or "Rescheduled")
// ======================================================
async function getFilteredLeads(sheets, sheetName) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:AA`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {
      const importantNote = row[10] ? row[10].trim() : "";
      const plannedDate = row[20] ? row[20].trim() : "";
      const actualDate = row[21] ? row[21].trim() : "";
      const status = row[22] ? row[22].trim() : "";
      const remarkFromT = row[19] ? row[19].trim() : "";
      const remarkFromY = row[25] ? row[25].trim() : "";
      let finalRemarks = "";
      if (remarkFromT) {
        finalRemarks = remarkFromT;
      } else if (remarkFromY) {
        finalRemarks = remarkFromY;
      }

      // Show only if status is empty or "Rescheduled" (case-insensitive)
      const showRow = !status || status.trim().toLowerCase() === "rescheduled";

      if (showRow && plannedDate) {
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

          // Yeh line add kar do – ab frontend ko followupCount milega
          followupCount: row[25] ? parseInt(row[25].trim(), 10) || 0 : 0,
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
// ROUTES - GET /list
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
// UPDATE - POST /update
// ======================================================
router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks, rescheduleDate } = req.body;
    console.log("📝 Update Request:", { sheetName, rowIndex, status, rescheduleDate, remarks });

    if (!sheetName || !rowIndex) {
      return res.status(400).json({ success: false, error: "Missing sheetName or rowIndex" });
    }

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // === Fetch current Followup Count from Z ===
    let currentFollowupCount = 0;
    try {
      const countResponse = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!Z${rowIndex}`,
      });

      const val = countResponse.data.values?.[0]?.[0];
      // Handle possible string formats like "5" or "5.0" or even time-like strings
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
      console.log(`Current Followup Count (Z${rowIndex}): ${currentFollowupCount}`);
    } catch (e) {
      console.warn(`⚠️ Could not read followup count from Z${rowIndex}:`, e.message);
      // Continue with 0 if read fails
    }

    const newFollowupCount = currentFollowupCount + 1;

    // === Always increment Followup Count in Z on EVERY update ===
    updates.push({
      range: `'${sheetName}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // === Main update logic ===
    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      // RESCHEDULE
      console.log("→ Processing RESCHEDULE");
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      updates.push({
        range: `'${sheetName}'!U${rowIndex}`,
        values: [[newPlannedDateTime]],
      });
      updates.push({
        range: `'${sheetName}'!W${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } else if (status === "Not Interested") {
      // NOT INTERESTED
      console.log("→ Processing NOT INTERESTED");
      updates.push({
        range: `'${sheetName}'!U${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${sheetName}'!W${rowIndex}`,
        values: [["Not Interested"]],
      });
    } else {
      // MARK DONE or any other status
      console.log("→ Processing DONE / STATUS UPDATE:", status || "Done");
      updates.push({
        range: `'${sheetName}'!U${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${sheetName}'!W${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // === Remarks (Y column) ===
    if (remarks && String(remarks).trim() !== "") {
      console.log("→ Updating remarks");
      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [[String(remarks).trim()]],
      });
    }

    // === Execute batch update ===
    if (updates.length > 0) {
      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: updates,
        },
      });
      console.log(`✅ Success: Row ${rowIndex} updated | New Followup Count = ${newFollowupCount}`);
    } else {
      console.log("No updates performed");
    }

    res.json({
      success: true,
      message: rescheduleDate
        ? "Visit Rescheduled successfully"
        : status === "Not Interested"
          ? "Marked as Not Interested"
          : "Field Visit marked as Done",
      newFollowupCount: newFollowupCount, // Frontend ko bata rahe hain
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