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

// Generates "DD/MM/YYYY HH:mm:ss" using current time + selected date
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const timePart = `${hours}:${minutes}:${seconds}`;

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-"); // Expecting YYYY-MM-DD
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  return `${formattedDate} ${timePart}`;
}

// --- READ DATA (GET) ---

async function getFilteredLeads(sheets, sheetName) {
  try {
    // Range A8:S tak (S = Latest Remarks)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:S`, 
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    rows.forEach((row, index) => {
      
      // --- COLUMN INDICES ---
      // L = 11 (Remark of Qualification - Initial)
      // M = 12, N = 13, O = 14
      // R = 17 (FollowUp Count)
      // S = 18 (Latest Remarks - Updated)

      const status = row[14] ? row[14].trim() : "";         // Col O
      const plannedDate = row[12] ? row[12].trim() : "";    // Col M
      const actualDate = row[13] ? row[13].trim() : "";     // Col N
      const followUpCountStr = row[17] ? row[17].trim() : "0"; // Col R
      
      const remarkL = row[11] ? row[11].trim() : "";        // Col L (Initial)
      const remarkS = row[19] ? row[19].trim() : "";        // Col T (Latest)

      // --- CONDITIONAL REMARKS LOGIC ---
      const countVal = parseInt(followUpCountStr) || 0;
      let finalRemarkToDisplay = "";

      if (countVal === 0) {
        // Agar count 0 hai, toh Column L dikhao
        finalRemarkToDisplay = remarkL;
      } else {
        // Agar count > 0 hai, toh Column S dikhao
        finalRemarkToDisplay = remarkS;
      }

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
          followUpCount: countVal,
          remarks: finalRemarkToDisplay, // Yahan conditional remark jayega
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
    } = req.body;

    console.log("📝 Updating NBDIN record:", { sheetName, rowIndex, status, remarks });

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields",
      });
    }

    const timestamp = new Date().toLocaleString("en-IN", {
      timeZone: "Asia/Kolkata",
      hour12: false,
    }).replace(",", "");

    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    // --- WRITING LOGIC (Always write to Column S for new remarks) ---

    // 1. Planned (Col M)
    let finalPlannedValue = "";
    if (nextFollowUpDate) {
      finalPlannedValue = getPlannedDateTime(nextFollowUpDate);
      updates.push({
        range: `'${sheetName}'!M${rowIndex}`,
        values: [[finalPlannedValue]],
      });
    }

    // 2. Actual (Col N)
    updates.push({
      range: `'${sheetName}'!N${rowIndex}`,
      values: [[timestamp]],
    });

    // 3. Status (Col O)
    updates.push({
      range: `'${sheetName}'!O${rowIndex}`,
      values: [[status]],
    });

    // 4. Field Visit (Col P)
    if (fieldVisitDate) {
      updates.push({
        range: `'${sheetName}'!P${rowIndex}`,
        values: [[fieldVisitDate]],
      });
    }

    // 5. Next FollowUp (Col Q)
    if (nextFollowUpDate) {
      updates.push({
        range: `'${sheetName}'!Q${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    }

    // 6. FollowUp Count (Col R)
    updates.push({
      range: `'${sheetName}'!R${rowIndex}`,
      values: [[newFollowUpCount.toString()]],
    });

    // 7. REMARKS (Always update Col S)
    if (remarks !== undefined) { 
      updates.push({
        range: `'${sheetName}'!S${rowIndex}`,
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

    console.log(`✅ Updated row ${rowIndex}. Count: ${newFollowUpCount}. Remarks logic handled.`);

    res.json({
      success: true,
      message: "Lead updated successfully",
      data: {
        sheetName,
        rowIndex,
        status,
        remarks,
        plannedDateTime: finalPlannedValue,
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