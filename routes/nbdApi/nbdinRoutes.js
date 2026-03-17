const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS"; // ✅ Only END USER sheet

// ✅ Lead Qualification Sheet (Source of truth for Doer)
const LEAD_QUAL_SPREADSHEET_ID = "17NsMDuq_woISO9CJTBh2e5BaZaKcSXkBEoEF6CNlDd0";
const LEAD_QUAL_SHEET_NAME = "FMS";

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

// ✅ Helper: Get doerTag from user info
// Maps logged-in user email to Doer column value
function getDoerTag(user) {
  if (!user) return null;

  // Admin sees all leads
  if (user.role === "admin" || user.assignedModule === "all") {
    return null; // null means no filter, show all
  }

  // Map email to doer tag
  const emailToDoerMap = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
  };

  return emailToDoerMap[user.email?.toLowerCase()] || null;
}

// --- READ DATA (GET) ---

async function getFilteredLeads(sheets, user) {
  try {
    // ✅ Extended range to A8:AT to include Doer column
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AT`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    // ✅ Get doer tag for current user
    const doerTag = getDoerTag(user);

    rows.forEach((row, index) => {
      // --- COLUMN MAPPING ---
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
      // AT [45] = Doer

      const importantNote = row[10] ? row[10].trim() : "";
      const status = row[14] ? row[14].trim() : "";
      const plannedDate = row[12] ? row[12].trim() : "";
      const actualDate = row[13] ? row[13].trim() : "";
      const followUpCountStr = row[17] ? row[17].trim() : "0";
      const pickAndDrop = row[18] ? row[18].trim() : "No";
      const oldRemarkL = row[11] ? row[11].trim() : "";
      const latestRemarkT = row[19] ? row[19].trim() : "";
      const doer = row[45] ? row[45].trim() : ""; // ✅ Doer column (AT = index 45)

      // --- REMARKS LOGIC ---
      const countVal = parseInt(followUpCountStr) || 0;
      let finalRemarkToDisplay = "";

      if (countVal === 0) {
        finalRemarkToDisplay = oldRemarkL;
      } else {
        finalRemarkToDisplay = latestRemarkT;
      }

      // --- STATUS FILTER ---
      if (status === "" || status === "No conversation" || status === "Next Follow Up") {
        // ✅ DOER FILTER: If doerTag exists, only show leads assigned to this user
        if (doerTag && doer !== doerTag) {
          return; // Skip this lead — not assigned to current user
        }

        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: NBD_SHEET_NAME,
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
          doer: doer, // ✅ Include doer in response
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching NBD leads:`, error.message);
    throw error;
  }
}

// --- GET ENDPOINT ---

router.get("/nbdin", async (req, res) => {
  try {
    console.log("📊 Fetching NBD Follow-up data (END USER only)...");
    console.log("🔍 req.user:", req.user ? JSON.stringify(req.user) : "❌ UNDEFINED");
    console.log("🔍 doerTag:", getDoerTag(req.user));
    console.log("🔍 Authorization header:", req.headers.authorization ? "✅ Present" : "❌ Missing");

    // ✅ Pass user info for doer-based filtering
    const leads = await getFilteredLeads(req.sheets, req.user);

    // Sort by planned date
    leads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });

    res.json({
      success: true,
      data: leads,
      total: leads.length,
    });
  } catch (error) {
    console.error("❌ Error fetching NBD leads:", error);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// --- UPDATE DATA (POST) ---

router.post("/nbdin/update", async (req, res) => {
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

    console.log("📝 Updating NBD Lead:", {
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
        range: `'${NBD_SHEET_NAME}'!M${rowIndex}`,
        values: [[finalPlannedValue]],
      });
    }

    // 2. Actual (Col N)
    updates.push({
      range: `'${NBD_SHEET_NAME}'!N${rowIndex}`,
      values: [[timestamp]],
    });

    // 3. Status (Col O)
    updates.push({
      range: `'${NBD_SHEET_NAME}'!O${rowIndex}`,
      values: [[status]],
    });

    // 4. Field Visit (Col P)
    if (fieldVisitDate) {
      updates.push({
        range: `'${NBD_SHEET_NAME}'!P${rowIndex}`,
        values: [[fieldVisitDate]],
      });
    }

    // 5. Next FollowUp (Col Q)
    if (nextFollowUpDate) {
      updates.push({
        range: `'${NBD_SHEET_NAME}'!Q${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    }

    // 6. FollowUp Count (Col R)
    updates.push({
      range: `'${NBD_SHEET_NAME}'!R${rowIndex}`,
      values: [[newFollowUpCount.toString()]],
    });

    // 7. Pick and Drop (Col S)
    if (pickAndDrop) {
      updates.push({
        range: `'${NBD_SHEET_NAME}'!S${rowIndex}`,
        values: [[pickAndDrop]],
      });
    }

    // 8. Remarks (Col T)
    if (remarks !== undefined) {
      updates.push({
        range: `'${NBD_SHEET_NAME}'!T${rowIndex}`,
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
      message: "NBD Lead updated successfully",
    });
  } catch (error) {
    console.error("❌ Error updating NBD lead:", error.message);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ============================================
// ✅ NEW: ASSIGN LEAD TO BDM (POST)
// Updates Doer column (V) in Lead Qualification Sheet
// ============================================

router.post("/nbdin/assign", async (req, res) => {
  try {
    const { uniqueId, assignTo } = req.body;

    console.log("🔄 Assigning lead:", { uniqueId, assignTo });

    if (!uniqueId || !assignTo) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields: uniqueId and assignTo",
      });
    }

    // Validate assignTo value
    const validAssignees = ["BDM1", "BDM2"];
    if (!validAssignees.includes(assignTo)) {
      return res.status(400).json({
        success: false,
        error: "Invalid assignTo value. Must be BDM1 or BDM2",
      });
    }

    // Step 1: Find the row in Lead Qualification Sheet by Unique ID
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: LEAD_QUAL_SPREADSHEET_ID,
      range: `'${LEAD_QUAL_SHEET_NAME}'!A:V`, // A to V (Doer column)
    });

    const rows = response.data.values || [];
    let targetRowIndex = -1;

    // Find the row with matching Unique ID (Column B = index 1)
    for (let i = 0; i < rows.length; i++) {
      const rowUniqueId = rows[i][1] ? rows[i][1].trim() : "";
      if (rowUniqueId === uniqueId) {
        targetRowIndex = i + 1; // Sheets are 1-indexed
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({
        success: false,
        error: `Lead with Unique ID "${uniqueId}" not found in Lead Qualification Sheet`,
      });
    }

    // Step 2: Update Doer column (V) in Lead Qualification Sheet
    await req.sheets.spreadsheets.values.update({
      spreadsheetId: LEAD_QUAL_SPREADSHEET_ID,
      range: `'${LEAD_QUAL_SHEET_NAME}'!V${targetRowIndex}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[assignTo]],
      },
    });

    console.log(`✅ Lead ${uniqueId} assigned to ${assignTo} at row ${targetRowIndex}`);

    res.json({
      success: true,
      message: `Lead ${uniqueId} assigned to ${assignTo} successfully`,
      data: {
        uniqueId,
        assignTo,
        rowIndex: targetRowIndex,
      },
    });
  } catch (error) {
    console.error("❌ Error assigning lead:", error.message);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

module.exports = router;