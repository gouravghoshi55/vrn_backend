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
 * Current timestamp generate karo
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
 * Date ko comparable format mein convert karo (sorting ke liye)
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
 * CONDITION: Column R (Planned) NOT NULL AND Column S (Actual) NULL
 * Row 8 se start
 * 
 * COLUMN STRUCTURE:
 * B=1 (Unique ID), C=2 (Customer Name), D=3 (Contact), 
 * E=4 (Interested In), F=5 (Project Selection), G=6 (Lead Source), 
 * H=7 (Lead Gen Number), I=8 (Lead Gen Name)
 * R=17 (Planned), S=18 (Actual), T=19 (Status), V=21 (Remarks)
 */
async function getFilteredLeads(sheets, sheetName) {
  try {
    // A8:V tak data fetch karo (Row 8 se start, Column A to V)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:V`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {
      // Column positions (0-indexed from A):
      // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8
      // R=17 (Planned), S=18 (Actual), T=19 (Status), V=21 (Remarks)
      
      const plannedDate = row[17] ? row[17].trim() : "";  // Column R (Planned)
      const actualDate = row[18] ? row[18].trim() : "";   // Column S (Actual)
      const status = row[19] ? row[19].trim() : "";       // Column T (Status)
      const remarks = row[21] ? row[21].trim() : "";      // Column V (Remarks)
      // CONDITION: Planned (R) NOT NULL AND Actual (S) NULL
      if (plannedDate && !actualDate) {
        filteredLeads.push({
          rowIndex: index + 8,                // Actual row number in sheet
          sheetName: sheetName,
          uniqueId: row[1] || "",             // Column B
          customerName: row[2] || "",         // Column C
          customerContact: row[3] || "",      // Column D
          interestedIn: row[4] || "",         // Column E
          projectSelection: row[5] || "",     // Column F
          leadSource: row[6] || "",           // Column G
          leadGenNumber: row[7] || "",        // Column H
          leadGenName: row[8] || "",          // Column I
          plannedDate: plannedDate,           // Column R
          status: status || "Pending",        // Column T
          remarks: remarks,                   // Column V
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
 * Dono sheets ka combined data - sorted by Planned Date (ascending)
 * CONDITION: Planned (R) NOT NULL AND Actual (S) NULL
 */
router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching Field Visit data...");
    console.log("   Condition: Planned (R) NOT NULL AND Actual (S) NULL");
    // Dono sheets se parallel fetch karo
    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);
    console.log(`   End User Leads: ${endUserLeads.length}`);
    console.log(`   Channel Partner Leads: ${channelPartnerLeads.length}`);
    // Combine both arrays
    let allLeads = [...endUserLeads, ...channelPartnerLeads];
    // Sort by Planned Date (ascending/increasing order)
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
 * Update Status (T), Remarks (V), and Actual timestamp (S)
 */
router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks } = req.body;
    console.log("📝 Updating Field Visit record:", { sheetName, rowIndex, status });
    // Validation
    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields: sheetName, rowIndex, status",
      });
    }
    // Current timestamp for Actual column (S)
    const timestamp = getCurrentTimestamp();
    // Batch update - S (Actual), T (Status), V (Remarks)
    const updates = [
      {
        range: `'${sheetName}'!S${rowIndex}`,
        values: [[timestamp]],
      },
      {
        range: `'${sheetName}'!T${rowIndex}`,
        values: [[status]],
      },
    ];
    // Add remarks if provided
    if (remarks) {
      updates.push({
        range: `'${sheetName}'!V${rowIndex}`,
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
    console.log(`✅ Updated Field Visit row ${rowIndex} in ${sheetName}`);
    console.log(`   Status: ${status}`);
    console.log(`   Remarks: ${remarks || "N/A"}`);
    console.log(`   Actual Timestamp: ${timestamp}`);
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
/**
 * DEBUG Route - Raw data dekhne ke liye
 * Browser: http://localhost:5000/api/field-visit/debug
 */
router.get("/debug", async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'END USER LEADS FMS'!A7:V10`,
    });
    res.json({
      rawData: response.data.values,
      message: "Row 7 = Headers, Row 8+ = Data",
      columnMapping: {
        "B (1)": "Unique ID",
        "C (2)": "Customer Name",
        "D (3)": "Customer Contact",
        "E (4)": "Interested In",
        "F (5)": "Project Selection",
        "G (6)": "Lead Source",
        "H (7)": "Lead Gen Number",
        "I (8)": "Lead Gen Name",
        "R (17)": "Planned (Field Visit)",
        "S (18)": "Actual (Field Visit)",
        "T (19)": "Status (Field Visit)",
        "V (21)": "Remarks",
      },
    });
  } catch (error) {
    res.json({ error: error.message });
  }
});
// ============================================
// Export Router
// ============================================
module.exports = router;