const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CP_SHEET_NAME = "Channel Partener Lead FMS";

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
function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";

  // CASE 1: datetime-local input (Example: "2026-02-14T11:43")
  if (dateStr.includes("T")) {
    const [datePart, timePart] = dateStr.split("T");
    const [year, month, day] = datePart.split("-");
    return `${day}/${month}/${year} ${timePart}:00`;
  }

  // CASE 2: Date only ("2026-02-14") -> Current Time add
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  const currentTime = `${hours}:${minutes}:${seconds}`;

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }

  return `${formattedDate} ${currentTime}`;
}

// --- READ DATA (GET) - Can Contact ---

async function getFilteredCPLeads(sheets, category) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${CP_SHEET_NAME}'!A8:AR`, // Extended to AR for Can Contact column
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      // --- COLUMN MAPPING (Same as NBD) ---
      // K [10] = Important Note
      // L [11] = OLD REMARKS (Initial/Old Remarks)
      // M [12] = Planned Date
      // N [13] = Actual Date
      // O [14] = Status
      // P [15] = Field Visit Date
      // Q [16] = Next FollowUp Date
      // R [17] = FollowUp Count
      // S [18] = Pick and Drop
      // T [19] = Latest Remarks
      // ...
      // AR [43] = Can Contact (Yes/No) ← NEW FOR CP

      const importantNote = row[10] ? row[10].trim() : "";
      const status = row[14] ? row[14].trim() : "";
      const plannedDate = row[12] ? row[12].trim() : "";
      const actualDate = row[13] ? row[13].trim() : "";
      const followUpCountStr = row[17] ? row[17].trim() : "0";
      const pickAndDrop = row[18] ? row[18].trim() : "No";
      const oldRemarkL = row[11] ? row[11].trim() : "";
      const latestRemarkT = row[19] ? row[19].trim() : "";
      
      const canContact = row[43] ? row[43].trim() : ""; // Column AR (Can Contact)

      // --- FILTER BY CATEGORY (Can Contact / Cannot Contact) ---
      const categoryMatch =
        (category === "can-contact" && canContact === "Yes") ||
        (category === "cannot-contact" && canContact === "No");

      if (!categoryMatch) return; // Skip if category doesn't match

      // --- REMARKS LOGIC ---
      const countVal = parseInt(followUpCountStr) || 0;
      let finalRemarkToDisplay = "";

      if (countVal === 0) {
        finalRemarkToDisplay = oldRemarkL;
      } else {
        finalRemarkToDisplay = latestRemarkT;
      }

      // --- STATUS FILTER (Same as NBD) ---
      if (status === "" || status === "No conversation" || status === "Next Follow Up") {
        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: CP_SHEET_NAME,
          uniqueId: row[1] || "",
          customerName: row[2] || "",
          customerContact: row[3] || "",
          interestedIn: row[4] || "",
          projectSelection: row[5] || "",
          leadSource: row[6] || "",
          leadGenNumber: row[7] || "",
          leadGenName: row[8] || "",
          importantNote: importantNote,
          pickAndDrop: pickAndDrop,
          plannedDate: plannedDate,
          actualDate: actualDate,
          status: status || "Pending",
          followUpCount: countVal,
          remarks: finalRemarkToDisplay,
          oldRemarks: oldRemarkL,
          canContact: canContact, // Send to frontend
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching CP ${category} leads:`, error.message);
    throw error;
  }
}

// --- GET ENDPOINTS ---

// Can Contact List
router.get("/can-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Can Contact Follow-up data...");
    
    const leads = await getFilteredCPLeads(req.sheets, "can-contact");

    leads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "can-contact",
    });
  } catch (error) {
    console.error("❌ Error fetching can-contact leads:", error);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// Cannot Contact List
router.get("/cannot-contact/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Cannot Contact Follow-up data...");
    
    const leads = await getFilteredCPLeads(req.sheets, "cannot-contact");

    leads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });

    res.json({
      success: true,
      data: leads,
      total: leads.length,
      category: "cannot-contact",
    });
  } catch (error) {
    console.error("❌ Error fetching cannot-contact leads:", error);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// --- UPDATE DATA (POST) ---

// Can Contact Update
router.post("/can-contact/update", async (req, res) => {
  try {
    const {
      rowIndex,
      status,
      fieldVisitDate,
      nextFollowUpDate,
      currentFollowUpCount,
      remarks,
      pickAndDrop,
    } = req.body;

    console.log("📝 Updating CP Can Contact:", {
      rowIndex,
      status,
      pickAndDrop,
      remarks,
    });

    if (!rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields",
      });
    }

    const timestamp = new Date()
      .toLocaleString("en-IN", {
        timeZone: "Asia/Kolkata",
        hour12: false,
      })
      .replace(",", "");

    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    // 1. Planned (Col M)
    if ((status === "No conversation" || status === "Next Follow Up") && nextFollowUpDate) {
      const finalPlannedValue = getPlannedDateTime(nextFollowUpDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!M${rowIndex}`,
        values: [[finalPlannedValue]],
      });
    }

    // 2. Actual (Col N)
    updates.push({
      range: `'${CP_SHEET_NAME}'!N${rowIndex}`,
      values: [[timestamp]],
    });

    // 3. Status (Col O)
    updates.push({
      range: `'${CP_SHEET_NAME}'!O${rowIndex}`,
      values: [[status]],
    });

    // 4. Field Visit (Col P)
    if (fieldVisitDate) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!P${rowIndex}`,
        values: [[fieldVisitDate]],
      });
    }

    // 5. Next FollowUp (Col Q)
    if (nextFollowUpDate) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!Q${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    }

    // 6. FollowUp Count (Col R)
    updates.push({
      range: `'${CP_SHEET_NAME}'!R${rowIndex}`,
      values: [[newFollowUpCount.toString()]],
    });

    // 7. Pick and Drop (Col S)
    if (pickAndDrop) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!S${rowIndex}`,
        values: [[pickAndDrop]],
      });
    }

    // 8. Remarks (Col T)
    if (remarks !== undefined) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!T${rowIndex}`,
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
      message: "CP Can Contact lead updated successfully",
    });
  } catch (error) {
    console.error("❌ Error updating CP Can Contact:", error.message);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// Cannot Contact Update
router.post("/cannot-contact/update", async (req, res) => {
  try {
    const {
      rowIndex,
      status,
      fieldVisitDate,
      nextFollowUpDate,
      currentFollowUpCount,
      remarks,
      pickAndDrop,
    } = req.body;

    console.log("📝 Updating CP Cannot Contact:", {
      rowIndex,
      status,
      pickAndDrop,
      remarks,
    });

    if (!rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields",
      });
    }

    const timestamp = new Date()
      .toLocaleString("en-IN", {
        timeZone: "Asia/Kolkata",
        hour12: false,
      })
      .replace(",", "");

    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    // Same update logic as Can Contact
    if ((status === "No conversation" || status === "Next Follow Up") && nextFollowUpDate) {
      const finalPlannedValue = getPlannedDateTime(nextFollowUpDate);
      updates.push({
        range: `'${CP_SHEET_NAME}'!M${rowIndex}`,
        values: [[finalPlannedValue]],
      });
    }

    updates.push({
      range: `'${CP_SHEET_NAME}'!N${rowIndex}`,
      values: [[timestamp]],
    });

    updates.push({
      range: `'${CP_SHEET_NAME}'!O${rowIndex}`,
      values: [[status]],
    });

    if (fieldVisitDate) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!P${rowIndex}`,
        values: [[fieldVisitDate]],
      });
    }

    if (nextFollowUpDate) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!Q${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    }

    updates.push({
      range: `'${CP_SHEET_NAME}'!R${rowIndex}`,
      values: [[newFollowUpCount.toString()]],
    });

    if (pickAndDrop) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!S${rowIndex}`,
        values: [[pickAndDrop]],
      });
    }

    if (remarks !== undefined) {
      updates.push({
        range: `'${CP_SHEET_NAME}'!T${rowIndex}`,
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
      message: "CP Cannot Contact lead updated successfully",
    });
  } catch (error) {
    console.error("❌ Error updating CP Cannot Contact:", error.message);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

module.exports = router;