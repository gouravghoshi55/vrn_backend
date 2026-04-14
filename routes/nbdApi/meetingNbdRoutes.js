const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

function getCurrentTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0"),
    m = String(now.getMonth() + 1).padStart(2, "0"),
    y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0"),
    mi = String(now.getMinutes()).padStart(2, "0"),
    s = String(now.getSeconds()).padStart(2, "0");
  return `${d}/${m}/${y} ${h}:${mi}:${s}`;
}

function getShortTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0"),
    mo = String(now.getMonth() + 1).padStart(2, "0"),
    y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0"),
    mi = String(now.getMinutes()).padStart(2, "0");
  return `[${d}/${mo}/${y} ${h}:${mi}]`;
}

function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  if (dateStr.includes(":") && dateStr.includes("/")) return dateStr.trim();
  if (dateStr.includes("T")) {
    const [d, t] = dateStr.split("T");
    const [y, m, dd] = d.split("-");
    const [hh, mm] = t.split(":");
    return `${dd}/${m}/${y} ${hh}:${mm}:00`;
  }
  if (dateStr.includes("-")) {
    const [y, m, d] = dateStr.split("-");
    return `${d}/${m}/${y} 00:00:00`;
  }
  return dateStr;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const p = dateStr.split(/[\/\-]/);
  if (p.length === 3) {
    if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]);
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return new Date(dateStr);
}
function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  if (user.assignedModule === "fsr") return null;
  const m = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
    "bdm3@company.com": "BDM3",
  };
  return m[user.email?.toLowerCase()] || null;
}
function getFSRDoerTag(user) {
  if (!user) return null;
  if (user.assignedModule !== "fsr") return null;
  const m = { "bdm4@company.com": "BDM4", "bdm5@company.com": "BDM5" };
  return m[user.email?.toLowerCase()] || null;
}

async function buildAppendedRemarks(sheets, sheetName, cellRange, newRemark) {
  if (!newRemark || !String(newRemark).trim()) return null;
  let existing = "";
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!${cellRange}`,
    });
    existing = r.data.values?.[0]?.[0] || "";
  } catch (e) {}
  const timestamped = `${getShortTimestamp()} ${String(newRemark).trim()}`;
  return existing.trim() ? `${existing.trim()}\n${timestamped}` : timestamped;
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
    const ts = getCurrentTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${LOGGER_SHEET_NAME}'!A:J`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [
          [
            ts,
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
  } catch (e) {
    console.error("❌ Logger failed:", e.message);
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
    const ts = getCurrentTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NOT_INTERESTED_SHEET}'!A:K`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [
          [
            ts,
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
  } catch (e) {
    console.error("❌ NI sheet failed:", e.message);
  }
}

async function getFilteredLeads(sheets, user) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${SHEET_NAME}'!A8:AN`,
  });
  const rows = response.data.values || [];
  const filtered = [];
  const doerTag = getDoerTag(user),
    fsrDoerTag = getFSRDoerTag(user);

  rows.forEach((row, index) => {
    const plannedDate = row[33] ? row[33].trim() : "";
    let status = row[35] ? row[35].trim() : "";
    const doer = row[38] ? row[38].trim() : "",
      fsrDoer = row[39] ? row[39].trim() : "";
    const oldRemarks = row[11] ? row[11].trim() : "",
      previousRemarksDate = row[13] ? row[13].trim() : "";
    const previousRemarks = row[19] ? row[19].trim() : "",
      latestOldRemarksDate = row[21] ? row[21].trim() : "";
    const latestOldRemarks = row[24] ? row[24].trim() : "",
      recentRemarksDate = row[27] ? row[27].trim() : "";
    const recentRemarks = row[32] ? row[32].trim() : "",
      currentRemarks = row[37] ? row[37].trim() : "";
    const displayRemarks =
      currentRemarks ||
      recentRemarks ||
      latestOldRemarks ||
      previousRemarks ||
      oldRemarks;

    const showRow =
      plannedDate &&
      (!status ||
        status.toLowerCase() === "rescheduled" ||
        status.toLowerCase() === "next field visit required");
    if (!showRow) return;
    if (doerTag && doer !== doerTag) return;
    if (fsrDoerTag && fsrDoer !== fsrDoerTag) return;
    if (!status.trim()) status = "Pending";

    filtered.push({
      rowIndex: index + 8,
      sheetName: SHEET_NAME,
      uniqueId: row[1] || "",
      customerName: row[2] || "",
      customerContact: row[3] || "",
      interestedIn: row[4] || "",
      projectSelection: row[5] || "",
      leadSource: row[6] || "",
      leadGenNumber: row[7] || "",
      leadGenName: row[8] || "",
      plannedDate,
      status,
      fsrDoer,
      remarks: displayRemarks,
      oldRemarks,
      previousRemarks,
      previousRemarksDate,
      latestOldRemarks,
      latestOldRemarksDate,
      recentRemarks,
      recentRemarksDate,
      doer,
    });
  });
  return filtered;
}

router.get("/list", async (req, res) => {
  try {
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

router.post("/update", async (req, res) => {
  try {
    const {
      rowIndex,
      status,
      rescheduleDate,
      nextFieldVisitDate,
      remarks,
      leadInfo,
      notInterestedReason,
    } = req.body;
    if (!rowIndex)
      return res
        .status(400)
        .json({ success: false, error: "Missing rowIndex" });

    const timestamp = getCurrentTimestamp();
    const updates = [];
    let currentFollowupCount = 0;
    try {
      const cr = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${SHEET_NAME}'!Z${rowIndex}`,
      });
      currentFollowupCount = parseInt(cr.data.values?.[0]?.[0]) || 0;
    } catch (e) {}
    const newFollowupCount = currentFollowupCount + 1;
    updates.push({
      range: `'${SHEET_NAME}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      updates.push({
        range: `'${SHEET_NAME}'!AH${rowIndex}`,
        values: [[getPlannedDateTime(rescheduleDate)]],
      });
      updates.push({
        range: `'${SHEET_NAME}'!AJ${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } else if (
      status === "Next Field Visit Required" &&
      nextFieldVisitDate &&
      String(nextFieldVisitDate).trim() !== ""
    ) {
      updates.push({
        range: `'${SHEET_NAME}'!AH${rowIndex}`,
        values: [[getPlannedDateTime(nextFieldVisitDate)]],
      });
      updates.push({
        range: `'${SHEET_NAME}'!AJ${rowIndex}`,
        values: [["Next Field Visit Required"]],
      });
    } else if (
      ["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(status)
    ) {
      updates.push({
        range: `'${SHEET_NAME}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${SHEET_NAME}'!AJ${rowIndex}`,
        values: [[status]],
      });
      updates.push({ range: `'${SHEET_NAME}'!AH${rowIndex}`, values: [[""]] });
      if (leadInfo)
        await appendToNotInterestedSheet(
          req.sheets,
          leadInfo,
          "Step 4 - Meeting",
          notInterestedReason || "",
          req.user?.email,
        );
    } else {
      updates.push({
        range: `'${SHEET_NAME}'!AI${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${SHEET_NAME}'!AJ${rowIndex}`,
        values: [[status || "Done"]],
      });
    }

    // ✅ Remarks — APPEND with timestamp
    if (remarks && String(remarks).trim() !== "") {
      const appended = await buildAppendedRemarks(
        req.sheets,
        SHEET_NAME,
        `AL${rowIndex}`,
        remarks,
      );
      if (appended)
        updates.push({
          range: `'${SHEET_NAME}'!AL${rowIndex}`,
          values: [[appended]],
        });
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });
    if (leadInfo)
      await appendToLogger(
        req.sheets,
        leadInfo,
        "Step 4 - Meeting",
        status || "Done",
        remarks || "",
        req.user?.email,
      );

    res.json({
      success: true,
      message: rescheduleDate
        ? "Rescheduled"
        : status === "Next Field Visit Required"
          ? "Next Field Visit Scheduled"
          : ["Not Interested", "Negotiation Failed", "Deal Not Done"].includes(
                status,
              )
            ? "Marked as " + status
            : "Updated",
      newFollowupCount,
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;
