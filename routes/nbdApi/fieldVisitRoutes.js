const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";

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

// ✅ Get FSR code from user email
function getFSRCode(user) {
  if (!user) return null;
  const fsrMap = {
    "bdm4@company.com": "BDM4",
    "bdm5@company.com": "BDM5",
  };
  return fsrMap[user.email?.toLowerCase()] || null;
}

// ✅ Get doerTag — AM column filter (BDM1/BDM2/BDM3 only)
function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  // FSR users — AM column se filter nahi, AN column se hoga
  if (user.assignedModule === "fsr") return null;
  const emailToDoerMap = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
    "bdm3@company.com": "BDM3",
  };
  return emailToDoerMap[user.email?.toLowerCase()] || null;
}

// ✅ Get FSR doer tag — AN column filter (BDM4/BDM5 only)
function getFSRDoerTag(user) {
  if (!user) return null;
  if (user.assignedModule !== "fsr") return null;
  const fsrMap = {
    "bdm4@company.com": "BDM4",
    "bdm5@company.com": "BDM5",
  };
  return fsrMap[user.email?.toLowerCase()] || null;
}

// ======================================================
// READ DATA - NBD Field Visit
// ======================================================

async function getFilteredLeads(sheets, user) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AN`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    const doerTag = getDoerTag(user);
    const fsrDoerTag = getFSRDoerTag(user);

    rows.forEach((row, index) => {
      const importantNote = row[10] ? row[10].trim() : "";
      const plannedDate = row[20] ? row[20].trim() : "";
      const actualDate = row[21] ? row[21].trim() : "";
      const status = row[22] ? row[22].trim() : "";
      const oldRemarks = row[11] ? row[11].trim() : "";
      const previousRemarksDate = row[13] ? row[13].trim() : "";
      const previousRemarks = row[19] ? row[19].trim() : "";
      const latestRemarks = row[24] ? row[24].trim() : "";
      const followupCount = row[25] ? parseInt(row[25].trim(), 10) || 0 : 0;
      const doer = row[38] ? row[38].trim() : "";   // AM = index 38
      const fsrDoer = row[39] ? row[39].trim() : ""; // AN = index 39

      let displayRemarks = latestRemarks || previousRemarks || oldRemarks;

      const showRow = !status || status.trim().toLowerCase() === "rescheduled";

      if (showRow && plannedDate) {
        // ✅ BDM1/BDM2/BDM3 filter — AM column
        if (doerTag && doer !== doerTag) return;

        // ✅ FSR filter — AN column
        // BDM4/BDM5 sirf wahi leads dekhen jahan AN mein unka code hai
        // Agar AN blank hai — FSR ko nahi dikhegi (abhi assign nahi hui)
        if (fsrDoerTag && fsrDoer !== fsrDoerTag) return;

        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: NBD_SHEET_NAME,
          uniqueId: row[1] || "",
          customerName: row[2] || "",
          customerContact: row[3] || "",
          interestedIn: row[4] || "",
          projectSelection: row[5] || "",
          leadSource: row[6] || "",
          leadGenNumber: row[7] || "",
          leadGenName: row[8] || "",
          importantNote,
          plannedDate,
          status: status || "Pending",
          followupCount,
          remarks: displayRemarks,
          oldRemarks,
          previousRemarks,
          previousRemarksDate,
          doer,
          fsrDoer, // ✅
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching NBD Field Visit leads:`, error.message);
    throw error;
  }
}

// ======================================================
// AUTO ASSIGN FSR — Load Balancing
// ======================================================

async function autoAssignFSR(sheets, rowIndex) {
  try {
    // Count current BDM4 and BDM5 assignments in AN column
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!AN8:AN`,
    });

    const rows = response.data.values || [];
    let bdm4Count = 0;
    let bdm5Count = 0;

    rows.forEach((row) => {
      const val = row[0] ? row[0].trim() : "";
      if (val === "BDM4") bdm4Count++;
      if (val === "BDM5") bdm5Count++;
    });

    // Assign to whoever has fewer leads
    const assignTo = bdm4Count <= bdm5Count ? "BDM4" : "BDM5";

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!AN${rowIndex}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[assignTo]] },
    });

    console.log(`✅ Auto-assigned to ${assignTo} (BDM4: ${bdm4Count}, BDM5: ${bdm5Count})`);
    return assignTo;
  } catch (error) {
    console.error("❌ Auto FSR assign failed:", error.message);
    return null;
  }
}

// ======================================================
// GET /list
// ======================================================

router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching NBD Field Visit data...");
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    console.error("❌ Error fetching NBD Field Visit:", error);
    res.status(500).json({ success: false, error: "Failed to fetch leads", message: error.message });
  }
});

// ======================================================
// POST /update
// ======================================================

router.post("/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate } = req.body;

    if (!rowIndex) {
      return res.status(400).json({ success: false, error: "Missing rowIndex" });
    }

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // Followup Count
    let currentFollowupCount = 0;
    try {
      const countResponse = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${NBD_SHEET_NAME}'!Z${rowIndex}`,
      });
      const val = countResponse.data.values?.[0]?.[0];
      currentFollowupCount = val ? parseInt(String(val).trim(), 10) || 0 : 0;
    } catch (e) {
      console.warn(`⚠️ Could not read followup count:`, e.message);
    }

    const newFollowupCount = currentFollowupCount + 1;
    updates.push({ range: `'${NBD_SHEET_NAME}'!Z${rowIndex}`, values: [[newFollowupCount]] });

    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      const newPlannedDateTime = getPlannedDateTime(rescheduleDate);
      updates.push({ range: `'${NBD_SHEET_NAME}'!U${rowIndex}`, values: [[newPlannedDateTime]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!W${rowIndex}`, values: [["Rescheduled"]] });
    } else if (status === "Not Interested") {
      updates.push({ range: `'${NBD_SHEET_NAME}'!V${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!W${rowIndex}`, values: [["Not Interested"]] });
    } else {
      // ✅ DONE — Auto assign FSR if AN column is blank
      updates.push({ range: `'${NBD_SHEET_NAME}'!V${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!W${rowIndex}`, values: [[status || "Done"]] });

      // Check if AN is already assigned
      let currentFSR = "";
      try {
        const fsrRes = await req.sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${NBD_SHEET_NAME}'!AN${rowIndex}`,
        });
        currentFSR = fsrRes.data.values?.[0]?.[0]?.trim() || "";
      } catch (e) {}

      // Auto assign only if AN is blank
      if (!currentFSR) {
        await autoAssignFSR(req.sheets, rowIndex);
      }
    }

    if (remarks && String(remarks).trim() !== "") {
      updates.push({ range: `'${NBD_SHEET_NAME}'!Y${rowIndex}`, values: [[String(remarks).trim()]] });
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    console.log(`✅ NBD Field Visit updated: Row ${rowIndex}`);

    res.json({
      success: true,
      message: rescheduleDate
        ? "Visit Rescheduled successfully"
        : status === "Not Interested"
        ? "Marked as Not Interested"
        : "Field Visit marked as Done",
      newFollowupCount,
    });
  } catch (error) {
    console.error("❌ NBD Field Visit update failed:", error.message);
    res.status(500).json({ success: false, error: "Failed to update lead", message: error.message });
  }
});

// ======================================================
// POST /migrate-fsr — ONE TIME — Purani blank leads assign
// ======================================================

router.post("/migrate-fsr", async (req, res) => {
  try {
    // Only admin
    if (req.user.role !== "admin" && req.user.assignedModule !== "all") {
      return res.status(403).json({ success: false, error: "Admin only" });
    }

    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AN`,
    });

    const rows = response.data.values || [];
    const updates = [];
    let bdm4Count = 0;
    let bdm5Count = 0;

    // Pehle existing counts nikalo
    rows.forEach((row) => {
      const an = row[39] ? row[39].trim() : "";
      if (an === "BDM4") bdm4Count++;
      if (an === "BDM5") bdm5Count++;
    });

    // Ab blank wali leads assign karo
    rows.forEach((row, index) => {
      const uniqueId = row[1] ? row[1].trim() : "";
      const an = row[39] ? row[39].trim() : "";

      if (uniqueId && !an) {
        const assignTo = bdm4Count <= bdm5Count ? "BDM4" : "BDM5";
        if (assignTo === "BDM4") bdm4Count++;
        else bdm5Count++;

        updates.push({
          range: `'${NBD_SHEET_NAME}'!AN${index + 8}`,
          values: [[assignTo]],
        });
      }
    });

    if (updates.length > 0) {
      // Batch mein 100-100 karo (Google Sheets limit)
      const batchSize = 100;
      for (let i = 0; i < updates.length; i += batchSize) {
        const batch = updates.slice(i, i + batchSize);
        await req.sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: { valueInputOption: "USER_ENTERED", data: batch },
        });
      }
    }

    res.json({
      success: true,
      message: `Migration complete`,
      assigned: updates.length,
      bdm4Total: bdm4Count,
      bdm5Total: bdm5Count,
    });
  } catch (error) {
    console.error("❌ Migration failed:", error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;