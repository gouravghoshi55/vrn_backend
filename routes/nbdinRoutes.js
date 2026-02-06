const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

async function getFilteredLeads(sheets, sheetName) {
  try {
    // A8:P tak data fetch karo (Row 8 se start, Column A to P)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:P`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {
      const status = row[12] ? row[12].trim() : ""; // Column M (Status)
      const plannedDate = row[10] ? row[10].trim() : ""; // Column K (Planned)
      const actualDate = row[11] ? row[11].trim() : ""; // Column L (Actual)
      const followUpCount = row[15] ? row[15].trim() : "0"; // Column P (FollowUp Count)

      const statusLower = status.toLowerCase();
      if (status === "" || statusLower === "no conversation") {
        filteredLeads.push({
          rowIndex: index + 8, // Actual row number in sheet
          sheetName: sheetName,
          uniqueId: row[1] || "", // Column B
          customerName: row[2] || "", // Column C
          customerContact: row[3] || "", // Column D
          interestedIn: row[4] || "", // Column E
          projectSelection: row[5] || "", // Column F
          leadSource: row[6] || "", // Column G
          leadGenNumber: row[7] || "", // Column H
          leadGenName: row[8] || "", // Column I
          plannedDate: plannedDate, // Column K
          actualDate: actualDate, // Column L
          status: status || "Pending", // Column M
          followUpCount: parseInt(followUpCount) || 0, // Column P
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
  const now = new Date();
  const day = String(now.getDate()).padStart(2, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const year = now.getFullYear();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
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
    if (nextFollowUpDate) {
      updates.push({
        range: `'${sheetName}'!K${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    }

    updates.push({
      range: `'${sheetName}'!L${rowIndex}`,
      values: [[timestamp]],
    });

    updates.push({
      range: `'${sheetName}'!M${rowIndex}`,
      values: [[status]],
    });

    if (fieldVisitDate) {
      updates.push({
        range: `'${sheetName}'!N${rowIndex}`,
        values: [[fieldVisitDate]],
      });
    }

    if (nextFollowUpDate) {
      updates.push({
        range: `'${sheetName}'!O${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    }

    updates.push({
      range: `'${sheetName}'!P${rowIndex}`,
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
