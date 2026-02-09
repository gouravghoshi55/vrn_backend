const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// Helper: Get filtered leads for Meeting
// Condition: Planned (AD) NOT NULL and Actual (AE) NULL
async function getFilteredLeads(sheets, sheetName) {
  try {
    // A8:AH tak data fetch karo (Row 8 se start)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:AH`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      // Column positions (0-indexed):
      // AD = 29 (Planned), AE = 30 (Actual), AF = 31 (Status), AH = 33 (Remarks)
      const plannedDate = row[29] ? row[29].trim() : "";  // Column AD
      const actualDate = row[30] ? row[30].trim() : "";   // Column AE
      const status = row[31] ? row[31].trim() : "";       // Column AF

      // Condition: Planned (AD) NOT NULL and Actual (AE) NULL
      // OR Status is empty or "Reschedule"
      if (plannedDate && !actualDate) {
        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: sheetName,
          uniqueId: row[1] || "",             // Column B
          customerName: row[2] || "",         // Column C
          customerContact: row[3] || "",      // Column D
          interestedIn: row[4] || "",         // Column E
          projectSelection: row[5] || "",     // Column F
          leadSource: row[6] || "",           // Column G
          leadGenNumber: row[7] || "",        // Column H
          leadGenName: row[8] || "",          // Column I
          plannedDate: plannedDate,           // Column AD
          status: status || "Pending",        // Column AF
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
}

// Parse date for sorting
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

// Get current timestamp
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

// ============================================
// API ROUTES
// ============================================

/**
 * GET /api/meeting-nbd/list
 * Dono sheets ka combined data - sorted by Planned Date (ascending)
 */
router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching Meeting NBD data...");
    console.log("   Condition: Planned (AD) NOT NULL and Actual (AE) NULL");

    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    console.log(`   End User Leads: ${endUserLeads.length}`);
    console.log(`   Channel Partner Leads: ${channelPartnerLeads.length}`);

    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    // Sort by Planned Date (ascending)
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
    console.error("❌ Error fetching Meeting NBD leads:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

/**
 * POST /api/meeting-nbd/update
 * Update Meeting status
 * 
 * Columns:
 * - AD (29): Planned - Gets rescheduled date if Reschedule
 * - AE (30): Actual - Gets timestamp when Done
 * - AF (31): Status - Done or Reschedule
 * - AH (33): Remarks
 */
router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, rescheduleDate, remarks } = req.body;

    console.log("📝 Updating Meeting NBD record:", { sheetName, rowIndex, status });

    // Validation
    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields: sheetName, rowIndex, status",
      });
    }

    if (status === "Reschedule" && !rescheduleDate) {
      return res.status(400).json({
        success: false,
        error: "Reschedule date is required when status is Reschedule",
      });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // If Reschedule: Update Planned (AD) with new date
    if (status === "Reschedule") {
      updates.push({
        range: `'${sheetName}'!AD${rowIndex}`,
        values: [[rescheduleDate]],
      });
    }

    // If Done: Update Actual (AE) with timestamp
    if (status === "Done") {
      updates.push({
        range: `'${sheetName}'!AE${rowIndex}`,
        values: [[timestamp]],
      });
    }

    // Update Status (AF)
    updates.push({
      range: `'${sheetName}'!AF${rowIndex}`,
      values: [[status]],
    });

    // Update Remarks (AH)
    if (remarks) {
      updates.push({
        range: `'${sheetName}'!AH${rowIndex}`,
        values: [[remarks]],
      });
    }

    // Batch update
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ Updated row ${rowIndex} in ${sheetName}`);
    console.log(`   Status: ${status}`);
    if (status === "Reschedule") console.log(`   Reschedule Date: ${rescheduleDate}`);
    if (status === "Done") console.log(`   Actual Timestamp: ${timestamp}`);
    if (remarks) console.log(`   Remarks: ${remarks}`);

    res.json({
      success: true,
      message: "Meeting updated successfully",
      data: {
        sheetName,
        rowIndex,
        status,
        rescheduleDate,
        remarks,
        actualTimestamp: status === "Done" ? timestamp : null,
      },
    });
  } catch (error) {
    console.error("❌ Error updating Meeting NBD lead:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

/**
 * GET /api/meeting-nbd/debug
 * Debug route to check raw data
 */
router.get("/debug", async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'END USER LEADS FMS'!A7:AH10`,
    });

    res.json({
      rawData: response.data.values,
      message: "Row 7 = Headers, Row 8+ = Data",
      columnMapping: {
        "AD (29)": "Planned",
        "AE (30)": "Actual",
        "AF (31)": "Status",
        "AH (33)": "Remarks",
      },
    });
  } catch (error) {
    res.json({ error: error.message });
  }
});

module.exports = router;