const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";
const NOT_INTERESTED_SHEET = "Not intrested reasons";

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

function formatDateToSheetStyle(dateInput) {
  if (!dateInput) return "";
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) return "";
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");
  const seconds = String(d.getSeconds()).padStart(2, "0");
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  if (user.assignedModule === "fsr") return null;
  const emailToDoerMap = { "bdm1@company.com": "BDM1", "bdm2@company.com": "BDM2", "bdm3@company.com": "BDM3" };
  return emailToDoerMap[user.email?.toLowerCase()] || null;
}

function getFSRDoerTag(user) {
  if (!user) return null;
  if (user.assignedModule !== "fsr") return null;
  const fsrMap = { "bdm4@company.com": "BDM4", "bdm5@company.com": "BDM5" };
  return fsrMap[user.email?.toLowerCase()] || null;
}

async function appendToLogger(sheets, lead, status, remarks, timestamp, stepName, userEmail) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${LOGGER_SHEET_NAME}'!A:J`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[timestamp, stepName, lead.uniqueId || "", lead.customerName || "", lead.customerContact || "", lead.interestedIn || "", lead.projectSelection || "", status, remarks || "", userEmail || ""]],
      },
    });
  } catch (error) {
    console.error("❌ Logger append failed:", error.message);
  }
}

// ✅ Not Interested Reasons sheet mein append
async function appendToNotInterestedSheet(sheets, leadInfo, stepName, reason, userEmail) {
  try {
    const timestamp = getCurrentTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NOT_INTERESTED_SHEET}'!A:K`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          timestamp, stepName,
          leadInfo.uniqueId || "", leadInfo.customerName || "", leadInfo.customerContact || "",
          leadInfo.interestedIn || "", leadInfo.projectSelection || "", leadInfo.leadSource || "",
          leadInfo.doer || "", reason || "", userEmail || "",
        ]],
      },
    });
    console.log(`✅ Not Interested reason logged for ${leadInfo.uniqueId}`);
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
    const doerTag = getDoerTag(user);
    const fsrDoerTag = getFSRDoerTag(user);

    return rows
      .map((row, index) => {
        const getCol = (idx) => (row[idx] ? String(row[idx]).trim() : "");
        const plannedDate = getCol(26);
        const actualDate = getCol(27);
        const status = getCol(28);
        const doer = getCol(38);
        const fsrDoer = getCol(39);

        const showRow = (plannedDate && !actualDate) || status === "No conversation" || status === "Next Follow Up" || status === "Next Field Visit Required";
        if (!showRow) return null;
        if (doerTag && doer !== doerTag) return null;
        if (fsrDoerTag && fsrDoer !== fsrDoerTag) return null;

        const oldRemarks = getCol(11);
        const previousRemarksDate = getCol(13);
        const previousRemarks = getCol(19);
        const latestOldRemarks = getCol(24);
        const latestOldRemarksDate = getCol(21);
        const currentRemarks = getCol(32);
        const displayRemarks = currentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

        return {
          rowIndex: index + 8, sheetName: NBD_SHEET_NAME,
          uniqueId: getCol(1), customerName: getCol(2), customerContact: getCol(3),
          interestedIn: getCol(4), projectSelection: getCol(5), leadSource: getCol(6),
          leadGenNumber: getCol(7), leadGenName: getCol(8), importantNote: getCol(10),
          plannedDate, status: status || "Pending", followUpCount: getCol(31) || "0",
          fsrDoer, remarks: displayRemarks, oldRemarks, previousRemarks, previousRemarksDate,
          latestOldRemarks, latestOldRemarksDate, doer,
        };
      })
      .filter(Boolean);
  } catch (error) {
    console.error("Error fetching NBD After Field Visit leads:", error.message);
    throw error;
  }
}

router.get("/list", async (req, res) => {
  try {
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => {
      const parse = (d) => { if (!d) return 0; const p = d.split(/[\/ :]/); return new Date(p[2], p[1] - 1, p[0], p[3] || 0, p[4] || 0).getTime(); };
      return parse(a.plannedDate) - parse(b.plannedDate);
    });
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    console.error("❌ Error fetching NBD After Field Visit:", error);
    res.status(500).json({ success: false, message: error.message });
  }
});

router.post("/update", async (req, res) => {
  try {
    const { rowIndex, status, remarks, rescheduleDate, dealMeetingDate, leadInfo, notInterestedReason } = req.body;

    if (!rowIndex) return res.status(400).json({ success: false, error: "Missing rowIndex" });

    let currentCount = 0;
    try {
      const countRes = await req.sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${NBD_SHEET_NAME}'!AF${rowIndex}` });
      currentCount = parseInt(countRes.data.values?.[0]?.[0]) || 0;
    } catch (e) {}

    const newCount = currentCount + 1;
    const updates = [];
    const timestamp = getCurrentTimestamp();

    updates.push({ range: `'${NBD_SHEET_NAME}'!AF${rowIndex}`, values: [[newCount]] });

    if (remarks && String(remarks).trim()) {
      updates.push({ range: `'${NBD_SHEET_NAME}'!AG${rowIndex}`, values: [[String(remarks).trim()]] });
    }

    if (status === "Not Interested") {
      updates.push({ range: `'${NBD_SHEET_NAME}'!AB${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, values: [["Not Interested"]] });

      // ✅ Log to Not Interested Reasons sheet
      if (leadInfo) {
        await appendToNotInterestedSheet(req.sheets, leadInfo, "Step 3 - After Field Visit", notInterestedReason || "", req.user?.email);
      }
    } else if (
      rescheduleDate && String(rescheduleDate).trim() !== "" &&
      ["No conversation", "Next Follow Up", "Next Field Visit Required"].includes(status || "")
    ) {
      const formatted = formatDateToSheetStyle(rescheduleDate);
      updates.push({ range: `'${NBD_SHEET_NAME}'!AA${rowIndex}`, values: [[formatted]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AE${rowIndex}`, values: [[formatted]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, values: [[status]] });
    } else {
      updates.push({ range: `'${NBD_SHEET_NAME}'!AB${rowIndex}`, values: [[timestamp]] });
      updates.push({ range: `'${NBD_SHEET_NAME}'!AC${rowIndex}`, values: [[status || "Done"]] });
      if (dealMeetingDate) {
        updates.push({ range: `'${NBD_SHEET_NAME}'!AD${rowIndex}`, values: [[formatDateToSheetStyle(dealMeetingDate)]] });
      }
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    if (status === "Next Field Visit Required" && leadInfo) {
      await appendToLogger(req.sheets, leadInfo, status, remarks, timestamp, "After Field Visit Follow-Up", req.user?.email);
    }

    res.json({ success: true, message: "Updated successfully", newFollowUpCount: newCount });
  } catch (error) {
    console.error("❌ NBD After Field Visit update failed:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;