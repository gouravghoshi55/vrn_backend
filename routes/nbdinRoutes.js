const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// ... (getFilteredLeads function same as before) ...
// Main copy paste kar raha hu getFilteredLeads taaki koi confusion na rahe range ko lekar
async function getFilteredLeads(sheets, sheetName) {
  try {
    // Range A8:R (FollowUp Count tak)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:R`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {
      // Adjusted Indices based on +2 shift
      const status = row[14] ? row[14].trim() : ""; // Col O
      const plannedDate = row[12] ? row[12].trim() : ""; // Col M
      const actualDate = row[13] ? row[13].trim() : ""; // Col N
      const followUpCount = row[17] ? row[17].trim() : "0"; // Col R

      const statusLower = status.toLowerCase();
      if (status === "" || statusLower === "no conversation") {
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
          plannedDate: plannedDate,
          actualDate: actualDate,
          status: status || "Pending",
          followUpCount: parseInt(followUpCount) || 0,
        });
      }
    });
    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
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

// 1. Current Timestamp for 'Actual' column
function getCurrentTimestamp() {
  const dt = new Date().toLocaleString("en-IN", {
    timeZone: "Asia/Kolkata",
    hour12: false,
  });
  return dt.replace(",", "");
}

// 2. NEW FUNCTION: Merges Selected Date with Current Time for 'Planned' column
function formatPlannedDateWithTime(dateInput) {
  if (!dateInput) return "";
  
  // dateInput usually comes as YYYY-MM-DD from frontend
  // We need to convert it to DD/MM/YYYY
  let datePart = dateInput;
  if (dateInput.includes("-")) {
      const parts = dateInput.split("-");
      // If format is YYYY-MM-DD
      if (parts[0].length === 4) {
          datePart = `${parts[2]}/${parts[1]}/${parts[0]}`;
      }
  }

  // Get current time only
  const now = new Date();
  const timePart = now.toLocaleTimeString("en-IN", {
    timeZone: "Asia/Kolkata",
    hour12: false,
  });

  // Return "DD/MM/YYYY HH:mm:ss"
  return `${datePart} ${timePart}`;
}

// ... (GET route remains same) ...
router.get("/nbdin", async (req, res) => {
    // ... (Your existing get logic) ...
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


// --- UPDATED POST ROUTE ---
router.post("/nbdin/update", async (req, res) => {
  try {
    const {
      sheetName,
      rowIndex,
      status,
      fieldVisitDate,
      nextFollowUpDate, // Input Format: YYYY-MM-DD
      currentFollowUpCount,
    } = req.body;

    console.log("📝 Updating NBDIN record:", { sheetName, rowIndex, status });

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields",
      });
    }

    const timestamp = getCurrentTimestamp(); // Current Date + Time
    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    // --- LOGIC FOR COLUMNS (Based on Screenshot +2 Shift) ---
    // Old Planned (K) -> New Planned (M)
    // Old Actual (L)  -> New Actual (N)
    // Old Status (M)  -> New Status (O)
    // Old Field (N)   -> New Field (P)
    // Old Next (O)    -> New Next (Q)
    // Old Count (P)   -> New Count (R)

    // 1. Column M: Planned (Next FollowUp Date + Current Time)
    if (nextFollowUpDate) {
      const plannedDateTime = formatPlannedDateWithTime(nextFollowUpDate);
      updates.push({
        range: `'${sheetName}'!M${rowIndex}`, 
        values: [[plannedDateTime]], // Example: "15/02/2026 15:52:33"
      });
    }

    // 2. Column N: Actual (Current Timestamp)
    updates.push({
      range: `'${sheetName}'!N${rowIndex}`, 
      values: [[timestamp]], 
    });

    // 3. Column O: Status
    updates.push({
      range: `'${sheetName}'!O${rowIndex}`, 
      values: [[status]], 
    });

    // 4. Column P: Field Visit Schedule Date
    if (fieldVisitDate) {
      updates.push({
        range: `'${sheetName}'!P${rowIndex}`, 
        values: [[fieldVisitDate]], 
      });
    }

    // 5. Column Q: Next FollowUp Date (Only Date usually)
    if (nextFollowUpDate) {
      updates.push({
        range: `'${sheetName}'!Q${rowIndex}`, 
        values: [[nextFollowUpDate]], 
      });
    }

    // 6. Column R: FollowUP Count
    updates.push({
      range: `'${sheetName}'!R${rowIndex}`, 
      values: [[newFollowUpCount.toString()]], 
    });

    // Execute Batch Update
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ Updated row ${rowIndex}. Planned (Col M) set to Date+Time.`);
    
    res.json({
      success: true,
      message: "Lead updated successfully",
      data: {
        sheetName,
        rowIndex,
        status,
        actualTimestamp: timestamp,
        followUpCount: newFollowUpCount,
      },
    });

  } catch (error) {
    console.error("❌ Error updating NBD IN lead:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

module.exports = router;