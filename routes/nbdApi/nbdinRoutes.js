const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const NOT_INTERESTED_SHEET = "Not intrested reasons";

const LEAD_QUAL_SPREADSHEET_ID = "17NsMDuq_woISO9CJTBh2e5BaZaKcSXkBEoEF6CNlDd0";
const LEAD_QUAL_SHEET_NAME = "FMS";

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  if (dateStr.includes("T")) {
    const [datePart, timePart] = dateStr.split("T");
    const [year, month, day] = datePart.split("-");
    return `${day}/${month}/${year} ${timePart}:00`;
  }
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  return `${formattedDate} ${hours}:${minutes}:${seconds}`;
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

function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  const emailToDoerMap = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
  };
  return emailToDoerMap[user.email?.toLowerCase()] || null;
}

// ✅ Not Interested Reasons sheet mein append
async function appendToNotInterestedSheet(sheets, leadRow, stepName, reason, userEmail) {
  try {
    const timestamp = getCurrentTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NOT_INTERESTED_SHEET}'!A:K`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          timestamp,                    // A - Timestamp
          stepName,                     // B - Step Name
          leadRow.uniqueId || "",       // C - Unique ID
          leadRow.customerName || "",   // D - Customer Name
          leadRow.customerContact || "",// E - Contact
          leadRow.interestedIn || "",   // F - Interested In
          leadRow.projectSelection || "",// G - Project
          leadRow.leadSource || "",     // H - Lead Source
          leadRow.doer || "",           // I - Doer (AM)
          reason || "",                 // J - Not Interested Reason
          userEmail || "",              // K - Updated By
        ]],
      },
    });
    console.log(`✅ Not Interested reason logged for ${leadRow.uniqueId}`);
  } catch (error) {
    console.error("❌ Not Interested sheet append failed:", error.message);
  }
}

async function getFilteredLeads(sheets, user) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AN`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];
    const doerTag = getDoerTag(user);

    rows.forEach((row, index) => {
      const importantNote = row[10] ? row[10].trim() : "";
      const status = row[14] ? row[14].trim() : "";
      const plannedDate = row[12] ? row[12].trim() : "";
      const actualDate = row[13] ? row[13].trim() : "";
      const followUpCountStr = row[17] ? row[17].trim() : "0";
      const pickAndDrop = row[18] ? row[18].trim() : "No";
      const oldRemarkL = row[11] ? row[11].trim() : "";
      const latestRemarkT = row[19] ? row[19].trim() : "";
      const doer = row[38] ? row[38].trim() : "";

      const countVal = parseInt(followUpCountStr) || 0;
      let finalRemarkToDisplay = countVal === 0 ? oldRemarkL : latestRemarkT;

      if (status === "" || status === "No conversation" || status === "Next Follow Up") {
        if (doerTag && doer !== doerTag) return;

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
          importantNote,
          pickAndDrop,
          plannedDate,
          actualDate,
          status: status || "Pending",
          followUpCount: countVal,
          remarks: finalRemarkToDisplay,
          oldRemarks: oldRemarkL,
          doer,
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching NBD leads:`, error.message);
    throw error;
  }
}

router.get("/nbdin", async (req, res) => {
  try {
    console.log("📊 Fetching NBD Follow-up data (END USER only)...");
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    console.error("❌ Error fetching NBD leads:", error);
    res.status(500).json({ success: false, error: "Failed to fetch leads", message: error.message });
  }
});

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
      notInterestedReason,
      leadInfo,
    } = req.body;

    if (!rowIndex || !status) {
      return res.status(400).json({ success: false, error: "Missing required fields" });
    }

    const timestamp = new Date()
      .toLocaleString("en-IN", { timeZone: "Asia/Kolkata", hour12: false })
      .replace(",", "");

    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    if ((status === "No conversation" || status === "Next Follow Up") && nextFollowUpDate) {
      const finalPlannedValue = getPlannedDateTime(nextFollowUpDate);
      updates.push({ range: `'${NBD_SHEET_NAME}'!M${rowIndex}`, values: [[finalPlannedValue]] });
    }

    updates.push({ range: `'${NBD_SHEET_NAME}'!N${rowIndex}`, values: [[timestamp]] });
    updates.push({ range: `'${NBD_SHEET_NAME}'!O${rowIndex}`, values: [[status]] });

    if (fieldVisitDate) {
      updates.push({ range: `'${NBD_SHEET_NAME}'!P${rowIndex}`, values: [[fieldVisitDate]] });
    }
    if (nextFollowUpDate) {
      updates.push({ range: `'${NBD_SHEET_NAME}'!Q${rowIndex}`, values: [[nextFollowUpDate]] });
    }

    updates.push({ range: `'${NBD_SHEET_NAME}'!R${rowIndex}`, values: [[newFollowUpCount.toString()]] });

    if (pickAndDrop) {
      updates.push({ range: `'${NBD_SHEET_NAME}'!S${rowIndex}`, values: [[pickAndDrop]] });
    }
    if (remarks !== undefined) {
      updates.push({ range: `'${NBD_SHEET_NAME}'!T${rowIndex}`, values: [[remarks]] });
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    // ✅ Not Interested → Log to sheet
    if (status === "Not Interested" && leadInfo) {
      await appendToNotInterestedSheet(
        req.sheets,
        leadInfo,
        "Step 1 - Follow Up",
        notInterestedReason || "",
        req.user?.email
      );
    }

    res.json({ success: true, message: "NBD Lead updated successfully" });
  } catch (error) {
    console.error("❌ Error updating NBD lead:", error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

router.post("/nbdin/assign", async (req, res) => {
  try {
    const { uniqueId, assignTo } = req.body;
    if (!uniqueId || !assignTo) {
      return res.status(400).json({ success: false, error: "Missing required fields: uniqueId and assignTo" });
    }
    const validAssignees = ["BDM1", "BDM2"];
    if (!validAssignees.includes(assignTo)) {
      return res.status(400).json({ success: false, error: "Invalid assignTo value. Must be BDM1 or BDM2" });
    }

    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: LEAD_QUAL_SPREADSHEET_ID,
      range: `'${LEAD_QUAL_SHEET_NAME}'!A:V`,
    });

    const rows = response.data.values || [];
    let targetRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const rowUniqueId = rows[i][1] ? rows[i][1].trim() : "";
      if (rowUniqueId === uniqueId) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ success: false, error: `Lead with Unique ID "${uniqueId}" not found` });
    }

    await req.sheets.spreadsheets.values.update({
      spreadsheetId: LEAD_QUAL_SPREADSHEET_ID,
      range: `'${LEAD_QUAL_SHEET_NAME}'!V${targetRowIndex}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[assignTo]] },
    });

    res.json({ success: true, message: `Lead ${uniqueId} assigned to ${assignTo} successfully`, data: { uniqueId, assignTo, rowIndex: targetRowIndex } });
  } catch (error) {
    console.error("❌ Error assigning lead:", error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;