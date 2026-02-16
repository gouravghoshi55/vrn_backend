const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// --- HELPER FUNCTIONS ---

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// Generates "DD/MM/YYYY HH:mm:ss"
// Handles both "YYYY-MM-DD" (adds current time) AND "YYYY-MM-DDTHH:mm" (uses selected time)
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";

  // CASE 1: Agar input "datetime-local" se aaya hai (Example: "2026-02-14T11:43")
  if (dateStr.includes("T")) {
    const [datePart, timePart] = dateStr.split("T"); // "2026-02-14" aur "11:43" alag ho gaye
    const [year, month, day] = datePart.split("-");

    // Return format: DD/MM/YYYY HH:mm:00
    return `${day}/${month}/${year} ${timePart}:00`;
  }

  // CASE 2: Agar input sirf Date hai ("2026-02-14") -> Current Time add karega
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const currentTime = `${hours}:${minutes}:${seconds}`;

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-"); // Expecting YYYY-MM-DD
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`; // Convert to DD/MM/YYYY
    }
  }

  return `${formattedDate} ${currentTime}`;
}

// --- READ DATA (GET) ---

async function getFilteredLeads(sheets, sheetName) {
  try {
    // Range badhakar T tak kar di hai
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:T`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {

      // --- COLUMN MAPPING ---
      // K [10] = Important Note
      // L [11] = OLD REMARKS (Initial/Old Remarks)
      // ...
      // R [17] = FollowUp Count
      // S [18] = Pick and Drop (NEW)
      // T [19] = Latest Remarks (SHIFTED)

      const importantNote = row[10] ? row[10].trim() : "";
      const status = row[14] ? row[14].trim() : "";
      const plannedDate = row[12] ? row[12].trim() : "";
      const actualDate = row[13] ? row[13].trim() : "";
      const followUpCountStr = row[17] ? row[17].trim() : "0";

      const pickAndDrop = row[18] ? row[18].trim() : "No"; // Read Col S

      const oldRemarkL = row[11] ? row[11].trim() : ""; // OLD Remark (Column L)
      const latestRemarkT = row[19] ? row[19].trim() : ""; // Latest Remark (Col T)

      // --- REMARKS LOGIC ---
      const countVal = parseInt(followUpCountStr) || 0;
      let finalRemarkToDisplay = "";

      if (countVal === 0) {
        finalRemarkToDisplay = oldRemarkL; // Pehli baar L dikhao
      } else {
        finalRemarkToDisplay = latestRemarkT; // Uske baad T dikhao
      }

      if (status === "" || status === "No conversation" || status === "Next Follow Up") {
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
          pickAndDrop: pickAndDrop, // Send to Frontend
          plannedDate: plannedDate,
          actualDate: actualDate,
          status: status || "Pending",
          followUpCount: countVal,
          remarks: finalRemarkToDisplay,
          oldRemarks: oldRemarkL, // ✅ OLD REMARKS bhej rahe hain frontend pe
        });
      }
    });
    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
}

router.get("/nbdin", async (req, res) => {
  try {
    console.log("📊 Fetching NBDIN data...");
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
    res.json({
      success: true,
      data: allLeads,
      total: allLeads.length,
    });
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch leads", message: error.message });
  }
});

// --- UPDATE DATA (POST) ---

router.post("/nbdin/update", async (req, res) => {
  try {
    const {
      sheetName,
      rowIndex,
      status,
      fieldVisitDate,
      nextFollowUpDate,
      currentFollowUpCount,
      remarks,
      pickAndDrop, // Frontend se "Yes" ya "No" aayega
    } = req.body;

    console.log("📝 Updating:", { sheetName, rowIndex, status, pickAndDrop, remarks });

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({ success: false, error: "Missing required fields" });
    }

    const timestamp = new Date().toLocaleString("en-IN", {
      timeZone: "Asia/Kolkata",
      hour12: false,
    }).replace(",", "");

    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    // --- WRITING LOGIC ---

    // 1. Planned (Col M)
    // ==================================================
    // 1. Planned (Col M)  -> No conversation OR Next Follow Up
    // ==================================================
    let finalPlannedValue = "";

    if (
      (status === "No conversation" || status === "Next Follow Up") &&
      nextFollowUpDate
    ) {
      finalPlannedValue = getPlannedDateTime(nextFollowUpDate);

      updates.push({
        range: `'${sheetName}'!M${rowIndex}`,
        values: [[finalPlannedValue]],
      });
    }

    // 2. Actual (Col N)
    updates.push({ range: `'${sheetName}'!N${rowIndex}`, values: [[timestamp]] });

    // 3. Status (Col O)
    updates.push({ range: `'${sheetName}'!O${rowIndex}`, values: [[status]] });

    // 4. Field Visit (Col P)
    if (fieldVisitDate) {
      updates.push({ range: `'${sheetName}'!P${rowIndex}`, values: [[fieldVisitDate]] });
    }

    // 5. Next FollowUp (Col Q)
    if (nextFollowUpDate) {
      updates.push({ range: `'${sheetName}'!Q${rowIndex}`, values: [[nextFollowUpDate]] });
    }

    // 6. FollowUp Count (Col R)
    updates.push({ range: `'${sheetName}'!R${rowIndex}`, values: [[newFollowUpCount.toString()]] });

    // ==================================================
    // 7. PICK AND DROP -> COLUMN S (Index 18)
    // ==================================================
    if (pickAndDrop) {
      updates.push({
        range: `'${sheetName}'!S${rowIndex}`,
        values: [[pickAndDrop]],
      });
    }

    // ==================================================
    // 8. REMARKS -> COLUMN T (Index 19)
    // ==================================================
    // Ye check karein ki remarks empty bhi ho toh bhi update ho ya nahi.
    // Usually undefined check kaafi hai.
    if (remarks !== undefined) {
      updates.push({
        range: `'${sheetName}'!T${rowIndex}`, // Yahan T kar diya hai
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
      message: "Lead updated successfully",
    });
  } catch (error) {
    console.error("❌ Error updating:", error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});
module.exports = router;