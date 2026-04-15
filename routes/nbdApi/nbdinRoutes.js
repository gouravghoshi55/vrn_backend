const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

const LEAD_QUAL_SPREADSHEET_ID = "17NsMDuq_woISO9CJTBh2e5BaZaKcSXkBEoEF6CNlDd0";
const LEAD_QUAL_SHEET_NAME = "FMS";

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4)
      return new Date(parts[0], parts[1] - 1, parts[2]);
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
  const h = String(now.getHours()).padStart(2, "0");
  const m = String(now.getMinutes()).padStart(2, "0");
  const s = String(now.getSeconds()).padStart(2, "0");
  let fd = dateStr;
  if (dateStr.includes("-")) {
    const p = dateStr.split("-");
    if (p[0].length === 4) fd = `${p[2]}/${p[1]}/${p[0]}`;
  }
  return `${fd} ${h}:${m}:${s}`;
}

function getCurrentTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0");
  const mo = String(now.getMonth() + 1).padStart(2, "0");
  const y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const s = String(now.getSeconds()).padStart(2, "0");
  return `${d}/${mo}/${y} ${h}:${mi}:${s}`;
}

// ✅ Short timestamp for remarks: [DD/MM/YYYY HH:MM]
function getShortTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0");
  const mo = String(now.getMonth() + 1).padStart(2, "0");
  const y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  return `[${d}/${mo}/${y} ${h}:${mi}]`;
}

function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  const emailToDoerMap = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
    "bdm6@company.com": "BDM6",
  };
  return emailToDoerMap[user.email?.toLowerCase()] || null;
}

// ✅ Read existing remarks from cell, prepend new remark with timestamp
async function buildAppendedRemarks(sheets, sheetName, cellRange, newRemark) {
  if (!newRemark || !String(newRemark).trim()) return null;
  let existingRemarks = "";
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!${cellRange}`,
    });
    existingRemarks = res.data.values?.[0]?.[0] || "";
  } catch (e) {}

  const timestamped = `${getShortTimestamp()} ${String(newRemark).trim()}`;
  if (existingRemarks.trim()) {
    return `${existingRemarks.trim()}\n${timestamped}`;
  }
  return timestamped;
}

async function appendToLogger(
  sheets,
  leadInfo,
  stepName,
  status,
  remarks,
  userEmail,
) {
  try {
    const timestamp = getCurrentTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${LOGGER_SHEET_NAME}'!A:J`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [
          [
            timestamp,
            stepName,
            leadInfo.uniqueId || "",
            leadInfo.customerName || "",
            leadInfo.customerContact || "",
            leadInfo.interestedIn || "",
            leadInfo.projectSelection || "",
            status,
            remarks || "",
            userEmail || "",
          ],
        ],
      },
    });
  } catch (error) {
    console.error("❌ Logger append failed:", error.message);
  }
}

async function appendToNotInterestedSheet(
  sheets,
  leadInfo,
  stepName,
  reason,
  userEmail,
) {
  try {
    const timestamp = getCurrentTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NOT_INTERESTED_SHEET}'!A:K`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [
          [
            timestamp,
            stepName,
            leadInfo.uniqueId || "",
            leadInfo.customerName || "",
            leadInfo.customerContact || "",
            leadInfo.interestedIn || "",
            leadInfo.projectSelection || "",
            leadInfo.leadSource || "",
            leadInfo.doer || "",
            reason || "",
            userEmail || "",
          ],
        ],
      },
    });
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

      if (
        status === "" ||
        status === "No conversation" ||
        status === "Next Follow Up"
      ) {
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
    throw error;
  }
}

router.get("/nbdin", async (req, res) => {
  try {
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    res
      .status(500)
      .json({
        success: false,
        error: "Failed to fetch leads",
        message: error.message,
      });
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
    if (!rowIndex || !status)
      return res
        .status(400)
        .json({ success: false, error: "Missing required fields" });

    const timestamp = new Date()
      .toLocaleString("en-IN", { timeZone: "Asia/Kolkata", hour12: false })
      .replace(",", "");
    const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;
    const updates = [];

    if (
      (status === "No conversation" || status === "Next Follow Up") &&
      nextFollowUpDate
    ) {
      updates.push({
        range: `'${NBD_SHEET_NAME}'!M${rowIndex}`,
        values: [[getPlannedDateTime(nextFollowUpDate)]],
      });
    }
    updates.push({
      range: `'${NBD_SHEET_NAME}'!N${rowIndex}`,
      values: [[timestamp]],
    });
    updates.push({
      range: `'${NBD_SHEET_NAME}'!O${rowIndex}`,
      values: [[status]],
    });
    if (fieldVisitDate)
      updates.push({
        range: `'${NBD_SHEET_NAME}'!P${rowIndex}`,
        values: [[fieldVisitDate]],
      });
    if (nextFollowUpDate)
      updates.push({
        range: `'${NBD_SHEET_NAME}'!Q${rowIndex}`,
        values: [[nextFollowUpDate]],
      });
    updates.push({
      range: `'${NBD_SHEET_NAME}'!R${rowIndex}`,
      values: [[newFollowUpCount.toString()]],
    });
    if (pickAndDrop)
      updates.push({
        range: `'${NBD_SHEET_NAME}'!S${rowIndex}`,
        values: [[pickAndDrop]],
      });

    // ✅ Remarks — APPEND with timestamp (new on top of old)
    if (remarks && String(remarks).trim() !== "") {
      const appendedRemarks = await buildAppendedRemarks(
        req.sheets,
        NBD_SHEET_NAME,
        `T${rowIndex}`,
        remarks,
      );
      if (appendedRemarks) {
        updates.push({
          range: `'${NBD_SHEET_NAME}'!T${rowIndex}`,
          values: [[appendedRemarks]],
        });
      }
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    // ✅ HAR status change pe Logger
    if (leadInfo)
      await appendToLogger(
        req.sheets,
        leadInfo,
        "Step 1 - Follow Up",
        status,
        remarks || "",
        req.user?.email,
      );
    if (status === "Not Interested" && leadInfo)
      await appendToNotInterestedSheet(
        req.sheets,
        leadInfo,
        "Step 1 - Follow Up",
        notInterestedReason || "",
        req.user?.email,
      );

    res.json({ success: true, message: "NBD Lead updated successfully" });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

router.post("/nbdin/assign", async (req, res) => {
  try {
    const { uniqueId, assignTo } = req.body;
    if (!uniqueId || !assignTo)
      return res.status(400).json({ success: false, error: "Missing fields" });
    if (!["BDM1", "BDM2", "BDM6"].includes(assignTo))
      return res
        .status(400)
        .json({ success: false, error: "Invalid assignTo" });

    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: LEAD_QUAL_SPREADSHEET_ID,
      range: `'${LEAD_QUAL_SHEET_NAME}'!A:V`,
    });
    const rows = response.data.values || [];
    let targetRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if ((rows[i][1] || "").trim() === uniqueId) {
        targetRowIndex = i + 1;
        break;
      }
    }
    if (targetRowIndex === -1)
      return res
        .status(404)
        .json({ success: false, error: `Lead "${uniqueId}" not found` });

    await req.sheets.spreadsheets.values.update({
      spreadsheetId: LEAD_QUAL_SPREADSHEET_ID,
      range: `'${LEAD_QUAL_SHEET_NAME}'!V${targetRowIndex}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[assignTo]] },
    });
    res.json({
      success: true,
      message: `Lead ${uniqueId} assigned to ${assignTo}`,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;
