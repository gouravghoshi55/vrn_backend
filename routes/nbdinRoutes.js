const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

async function getFilteredLeads(sheets, sheetName) {
  try {
    // UPDATED: Range increased to 'R' to cover FollowUp Count (Column R)
    // A8:R tak data fetch karo (Row 8 se start, Column A to R)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:R`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {
      // UPDATED COLUMN INDICES BASED ON SCREENSHOT
      // A=0, B=1, ... M=12, N=13, O=14, R=17
      
      const status = row[14] ? row[14].trim() : ""; // Column O (Status) - New Index 14
      const plannedDate = row[12] ? row[12].trim() : ""; // Column M (Planned) - New Index 12
      const actualDate = row[13] ? row[13].trim() : ""; // Column N (Actual) - New Index 13
      const followUpCount = row[17] ? row[17].trim() : "0"; // Column R (FollowUp Count) - New Index 17

      const statusLower = status.toLowerCase();
      if (status === "" || statusLower === "no conversation") {
        filteredLeads.push({
          rowIndex: index + 8, // Actual row number in sheet
          sheetName: sheetName,
          uniqueId: row[1] || "", // Column B (Unique ID)
          customerName: row[2] || "", // Column C (Customer Name)
          customerContact: row[3] || "", // Column D (Contact)
          interestedIn: row[4] || "", // Column E
          projectSelection: row[5] || "", // Column F
          leadSource: row[6] || "", // Column G
          leadGenNumber: row[7] || "", // Column H
          leadGenName: row[8] || "", // Column I
          plannedDate: plannedDate, // Column M
          actualDate: actualDate, // Column N
          status: status || "Pending", // Column O
          followUpCount: parseInt(followUpCount) || 0, // Column R
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
    if (parts[0].length === 4) {
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

function getCurrentTimestamp() {
  const dt = new Date().toLocaleString("en-IN", {
    timeZone: "Asia/Kolkata",
    hour12: false,
  });
  return dt.replace(",", "");
}

router.get("/nbdin", async (req, res) => {
  try {
    console.log("📊 Fetching NBDIN data...");
    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);
    console.log(`   End User Leads: ${endUserLeads.length}`);
    console.log(`   Channel Partner Leads: ${channelPartnerLeads.length}`);
    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    allLeads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });
    console.log(`✅ Total Leads: ${allLeads.length}`);
    res.json({
      success: true,
      data: allLeads,
      total: allLeads.length,
    });
  } catch (error) {
    console.error("❌ Error fetching NBD IN leads:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

router.post("/nbdin/update", async (req, res) => {
  try {
    const {
      sheetName,
      rowIndex,
      status,
      fieldVisitDate,
      nextFollowUpDate,
      currentFollowUpCount,
    } = req.body;
    console.log("📝 Updating NBDIN record:", { sheetName, rowIndex, status });
    
    // Validation
    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields: sheetName, rowIndex, status",
      });
    }

    // Current timestamp for Actual column
    const timestamp = getCurrentTimestamp();
    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    // --- UPDATED WRITING LOGIC BASED ON NEW COLUMNS ---

    // 1. Update 'Planned' (Column M) with Next FollowUp Date
    if (nextFollowUpDate) {
      updates.push({
        range: `'${sheetName}'!M${rowIndex}`, // Changed from K to M
        values: [[nextFollowUpDate]],
      });
    }

    // 2. Update 'Actual' (Column N) with Timestamp
    updates.push({
      range: `'${sheetName}'!N${rowIndex}`, // Changed from L to N
      values: [[timestamp]],
    });

    // 3. Update 'Status' (Column O)
    updates.push({
      range: `'${sheetName}'!O${rowIndex}`, // Changed from M to O
      values: [[status]],
    });

    // 4. Update 'Field Visit Schedule date' (Column P)
    if (fieldVisitDate) {
      updates.push({
        range: `'${sheetName}'!P${rowIndex}`, // Changed from N to P
        values: [[fieldVisitDate]],
      });
    }

    // 5. Update 'Next FollowUP Date' (Column Q)
    if (nextFollowUpDate) {
      updates.push({
        range: `'${sheetName}'!Q${rowIndex}`, // Changed from O to Q
        values: [[nextFollowUpDate]],
      });
    }

    // 6. Update 'FollowUP Count' (Column R)
    updates.push({
      range: `'${sheetName}'!R${rowIndex}`, // Changed from P to R
      values: [[newFollowUpCount.toString()]],
    });

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ Updated row ${rowIndex} in ${sheetName}`);
    console.log(`   Status: ${status}`);
    console.log(`   Field Visit Date: ${fieldVisitDate || "N/A"}`);
    console.log(`   Next FollowUp Date: ${nextFollowUpDate || "N/A"}`);
    console.log(`   Actual Timestamp: ${timestamp}`);
    console.log(`   FollowUp Count: ${newFollowUpCount}`);
    
    res.json({
      success: true,
      message: "Lead updated successfully",
      data: {
        sheetName,
        rowIndex,
        status,
        fieldVisitDate,
        nextFollowUpDate,
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