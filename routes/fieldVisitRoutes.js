const express = require("express");
const router = express.Router();

// ============================================
// Configuration
// ============================================
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// ============================================
// Helper Functions
// ============================================

/**
 * Current timestamp generate karo (DD/MM/YYYY HH:mm:ss)
 */
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

/**
 * Date parse helper
 */
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

/**
 * Sheet se filtered data lao
 * CONDITION: Planned (T) is present AND Actual (U) is empty
 * 
 * COLUMN MAPPING (Based on Screenshot):
 * T (19) = Planned
 * U (20) = Actual
 * V (21) = Status
 * X (23) = Remarks
 */
async function getFilteredLeads(sheets, sheetName) {
  try {
    // UPDATED: Range increased to 'X' to cover Remarks
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:X`,
    });
    
    const rows = response.data.values || [];
    const filteredLeads = [];
    
    rows.forEach((row, index) => {
      // Corrected Indices:
      const plannedDate = row[19] ? row[19].trim() : "";  // Column T (Index 19)
      const actualDate = row[20] ? row[20].trim() : "";   // Column U (Index 20)
      const status = row[21] ? row[21].trim() : "";       // Column V (Index 21)
      const remarks = row[23] ? row[23].trim() : "";      // Column X (Index 23)

      // LOGIC: Show lead if Planned exists AND (Actual is empty OR Status is Pending)
      // Usually for Step 3 list: plannedDate && !actualDate
      if (plannedDate && !actualDate) {
        filteredLeads.push({
          rowIndex: index + 8,                // Actual row number
          sheetName: sheetName,
          uniqueId: row[1] || "",             // Column B
          customerName: row[2] || "",         // Column C
          customerContact: row[3] || "",      // Column D
          interestedIn: row[4] || "",         // Column E
          projectSelection: row[5] || "",     // Column F
          leadSource: row[6] || "",           // Column G
          leadGenNumber: row[7] || "",        // Column H
          leadGenName: row[8] || "",          // Column I
          plannedDate: plannedDate,           // Column T
          status: status || "Pending",        // Column V
          remarks: remarks,                   // Column X
        });
      }
    });
    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
}

// ============================================
// API ROUTES
// ============================================

/**
 * GET /api/field-visit/list
 */
router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching Field Visit (Step 3) data...");
    
    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    // Sort by Planned Date
    allLeads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });

    console.log(`✅ Total Field Visit Leads: ${allLeads.length}`);
    res.json({
      success: true,
      data: allLeads,
      total: allLeads.length,
    });
  } catch (error) {
    console.error("❌ Error fetching Field Visit leads:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

/**
 * POST /api/field-visit/update
 * Updates:
 * - Actual Date -> Column U
 * - Status -> Column V
 * - Remarks -> Column X
 */
router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks } = req.body;
    console.log("📝 Updating Field Visit record:", { sheetName, rowIndex, status });

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields",
      });
    }

    const timestamp = getCurrentTimestamp();
    
    // UPDATED WRITING LOGIC
    const updates = [
      // 1. Update Actual (Column U)
      {
        range: `'${sheetName}'!U${rowIndex}`,
        values: [[timestamp]],
      },
      // 2. Update Status (Column V)
      {
        range: `'${sheetName}'!V${rowIndex}`,
        values: [[status]],
      },
      // 3. Update Remarks (Column X)
      {
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[remarks || ""]],
      }
    ];

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ Updated Field Visit row ${rowIndex} in ${sheetName}`);
    console.log(`   Actual (Col U): ${timestamp}`);
    console.log(`   Status (Col V): ${status}`);
    console.log(`   Remarks (Col X): ${remarks}`);

    res.json({
      success: true,
      message: "Field Visit updated successfully",
      data: {
        sheetName,
        rowIndex,
        status,
        remarks,
        actualTimestamp: timestamp,
      },
    });
  } catch (error) {
    console.error("❌ Error updating Field Visit lead:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

// Debug Route to verify columns
router.get("/debug", async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEETS.END_USER}'!T7:X8`, // Fetch headers for checking
    });
    res.json({
      headers: response.data.values,
      mappingCheck: "Col T should be Planned, U Actual, V Status, X Remarks"
    });
  } catch (error) {
    res.json(error);
  }
});

module.exports = router;