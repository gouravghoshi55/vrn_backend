const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// --- HELPER FUNCTIONS ---

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

// Date Formatter
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const timePart = `${hours}:${minutes}:${seconds}`;

  let formattedDate = dateStr;
  // Agar date YYYY-MM-DD format mein aayi hai toh usse DD/MM/YYYY convert karein
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-"); 
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  // Google Sheets mein consistent format ke liye Time add kar rahe hain
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

// --- READ DATA ---
async function getFilteredLeads(sheets, sheetName) {
  try {
    // Column AF (Index 31) tak data read kar rahe hain
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:AF`, 
    });
    
    const rows = response.data.values || [];
    const filteredLeads = [];
    
    rows.forEach((row, index) => {
      // --- COLUMN MAPPING ---
      // Z  = 25 (Planned)
      // AA = 26 (Actual)
      // AB = 27 (Status)
      // AD = 29 (Next FollowUP Date)
      // AE = 30 (FollowUP Count)
      // AF = 31 (Remarks)

      const plannedDate = row[25] ? row[25].trim() : "";
      const actualDate = row[26] ? row[26].trim() : "";
      const status = row[27] ? row[27].trim() : "";
      const remarks = row[31] ? row[31].trim() : ""; 

      const isPendingVisit = plannedDate && !actualDate;
      const isNoConversation = status === "No conversation";

      if (isPendingVisit || isNoConversation) {
        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: sheetName,
          uniqueId: row[1] || "",
          customerName: row[2] || "",
          plannedDate: plannedDate,
          status: status || "Pending",
          remarks: remarks,
          // Extra info agar frontend pe chahiye ho
          followUpCount: row[30] || "0" 
        });
      }
    });
    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
}

// --- ROUTES ---

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
    res.status(500).json({ success: false, error: "Failed to fetch leads", message: error.message });
  }
});

// --- UPDATE ROUTE (FIXED LOGIC) ---
router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks, rescheduleDate } = req.body;
    console.log("📝 Updating Data:", { sheetName, rowIndex, status, rescheduleDate });

    if (!sheetName || !rowIndex) {
      return res.status(400).json({ success: false, error: "Missing required fields" });
    }

    // STEP 1: Current FollowUP Count (Column AE) read karein
    const countResponse = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!AE${rowIndex}`,
    });

    let currentCount = 0;
    if (countResponse.data.values && countResponse.data.values[0]) {
      currentCount = parseInt(countResponse.data.values[0][0]) || 0;
    }
    const newCount = currentCount + 1; // Count increase by 1

    const updates = [];

    // --- LOGIC: ALWAYS UPDATE COUNT & REMARKS ---
    
    // 1. Update FollowUP Count (Column AE)
    updates.push({
      range: `'${sheetName}'!AE${rowIndex}`,
      values: [[newCount]],
    });

    // 2. Update Remarks (Column AF)
    if (remarks !== undefined) {
      updates.push({
        range: `'${sheetName}'!AF${rowIndex}`,
        values: [[remarks]],
      });
    }

    // --- LOGIC: RESCHEDULE VS ACTUAL UPDATE ---

    if (rescheduleDate) {
      // SCENARIO 1: RESCHEDULE (Date Override)
      const formattedDate = getPlannedDateTime(rescheduleDate);

      // A. Update Next FollowUP Date (Column AD)
      updates.push({
        range: `'${sheetName}'!AD${rowIndex}`,
        values: [[formattedDate]],
      });

      // B. Override Planned Date (Column Z) with same date
      updates.push({
        range: `'${sheetName}'!Z${rowIndex}`,
        values: [[formattedDate]],
      });

      // (Optional) Agar reschedule ho raha hai toh Status bhi update kar sakte hain
      if (status) {
         updates.push({
          range: `'${sheetName}'!AB${rowIndex}`, // Column AB (Status)
          values: [[status]],
        });
      }

    } else {
      // SCENARIO 2: VISIT DONE / STATUS UPDATE (No Date Change)
      const timestamp = getCurrentTimestamp();

      // A. Update Actual Time (Column AA)
      updates.push({
        range: `'${sheetName}'!AA${rowIndex}`,
        values: [[timestamp]],
      });

      // B. Update Status (Column AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[status]],
      });
    }

    // Execute Batch Update
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    console.log(`✅ Row ${rowIndex} Updated. New Count: ${newCount}`);

    res.json({
      success: true,
      message: "Lead updated successfully",
      newFollowUpCount: newCount
    });

  } catch (error) {
    console.error("❌ Error updating lead:", error.message);
    res.status(500).json({ success: false, error: "Failed to update lead", message: error.message });
  }
});

module.exports = router;