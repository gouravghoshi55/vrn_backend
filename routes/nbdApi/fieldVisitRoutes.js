const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

function getCurrentTimestamp() {
  const now = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
  const d = String(now.getUTCDate()).padStart(2, "0"),
    mo = String(now.getUTCMonth() + 1).padStart(2, "0"),
    y = now.getUTCFullYear();
  const h = String(now.getUTCHours()).padStart(2, "0"),
    mi = String(now.getUTCMinutes()).padStart(2, "0"),
    s = String(now.getUTCSeconds()).padStart(2, "0");
  return `${d}/${mo}/${y} ${h}:${mi}:${s}`;
}

function getShortTimestamp() {
  const now = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
  const d = String(now.getUTCDate()).padStart(2, "0"),
    mo = String(now.getUTCMonth() + 1).padStart(2, "0"),
    y = now.getUTCFullYear();
  const h = String(now.getUTCHours()).padStart(2, "0"),
    mi = String(now.getUTCMinutes()).padStart(2, "0");
  return `[${d}/${mo}/${y} ${h}:${mi}]`;
}

function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  if (dateStr.includes("T")) {
    const [dt, t] = dateStr.split("T");
    const [y, m, d] = dt.split("-");
    return `${d}/${m}/${y} ${t}:00`;
  }
  const now = new Date();
  let fd = dateStr;
  if (dateStr.includes("-")) {
    const p = dateStr.split("-");
    if (p[0].length === 4) fd = `${p[2]}/${p[1]}/${p[0]}`;
  }
  return `${fd} ${String(now.getHours()).padStart(2, "0")}:${String(
    now.getMinutes(),
  ).padStart(2, "0")}:${String(now.getSeconds()).padStart(2, "0")}`;
}

function addDaysToSheetDate(sheetDateStr, days) {
  let baseDate;
  if (!sheetDateStr || !String(sheetDateStr).trim()) {
    baseDate = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
  } else {
    const str = String(sheetDateStr).trim();
    const parts = str.split(" ");
    const datePart = parts[0];
    const timePart = parts[1] || "10:00:00";
    const dp = datePart.split("/");
    if (dp.length !== 3) {
      baseDate = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
    } else {
      const [d, m, y] = dp;
      const [hh = "10", mm = "00", ss = "00"] = timePart.split(":");
      baseDate = new Date(
        parseInt(y, 10),
        parseInt(m, 10) - 1,
        parseInt(d, 10),
        parseInt(hh, 10),
        parseInt(mm, 10),
        parseInt(ss, 10),
      );
      if (isNaN(baseDate.getTime()))
        baseDate = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
    }
  }
  baseDate.setDate(baseDate.getDate() + days);
  const dd = String(baseDate.getDate()).padStart(2, "0");
  const mo = String(baseDate.getMonth() + 1).padStart(2, "0");
  const yy = baseDate.getFullYear();
  const hh = String(baseDate.getHours()).padStart(2, "0");
  const mi = String(baseDate.getMinutes()).padStart(2, "0");
  const ss = String(baseDate.getSeconds()).padStart(2, "0");
  return `${dd}/${mo}/${yy} ${hh}:${mi}:${ss}`;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const p = dateStr.split(/[/-]/);
  if (p.length === 3) {
    if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]);
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return new Date(dateStr);
}

function getFSRCode(user) {
  if (!user) return null;
  const m = {
    "bdm4@company.com": "BDM4",
    "bdm5@company.com": "BDM5",
    "varun@company.com": "Varun Sir",
    "bdm7@company.com": "BDM7",
  };
  return m[user.email?.toLowerCase()] || null;
}

function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  if (user.assignedModule === "fsr") return null;
  const m = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
    "bdm3@company.com": "BDM3",
    "bdm6@company.com": "BDM6",
    "varun@company.com": "Varun Sir",
    "mohan@company.com": "Mohan Sir",
  };
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
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AN`,
    });
    const rows = response.data.values || [];
    const filteredLeads = [];
    const doerTag = getDoerTag(user);

    rows.forEach((row, index) => {
      const plannedDate = row[20] ? row[20].trim() : "";
      const status = row[22] ? row[22].trim() : "";
      const oldRemarks = row[11] ? row[11].trim() : "";
      const previousRemarksDate = row[13] ? row[13].trim() : "";
      const previousRemarks = row[19] ? row[19].trim() : "";
      const latestRemarks = row[24] ? row[24].trim() : "";
      const followupCount = row[25] ? parseInt(row[25].trim(), 10) || 0 : 0;
      const doer = row[38] ? row[38].trim() : "";

      // ✅ FIX: CNP leads को Field Visit से exclude करो
      // CNP leads सिर्फ CNP page में दिखेंगी
      const showRow =
        !status ||
        status.trim().toLowerCase() === "rescheduled" ||
        status.trim().toLowerCase() === "next followup required" ||
        status.trim().toLowerCase() === "cold";

      if (showRow && plannedDate) {
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
          importantNote: row[10] ? row[10].trim() : "",
          plannedDate,
          status: status || "Pending",
          followupCount,
          remarks: latestRemarks || previousRemarks || oldRemarks,
          oldRemarks,
          previousRemarks,
          previousRemarksDate,
          doer,
        });
      }
    });
    return filteredLeads;
  } catch (error) {
    throw error;
  }
}

router.get("/list", async (req, res) => {
  try {
    const leads = await getFilteredLeads(req.sheets, req.user);
    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    res
      .status(500)
      .json({ success: false, error: "Failed", message: error.message });
  }
});

router.post("/update", async (req, res) => {
  try {
    const {
      rowIndex,
      status,
      remarks,
      rescheduleDate,
      notInterestedReason,
      leadInfo,
      // ✅ FIX: actionType frontend से भेजो ताकि backend को पता चले
      actionType,
    } = req.body;

    if (!rowIndex)
      return res
        .status(400)
        .json({ success: false, error: "Missing rowIndex" });

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // ✅ Followup Count increment
    let currentFollowupCount = 0;
    try {
      const cr = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${NBD_SHEET_NAME}'!Z${rowIndex}`,
      });
      currentFollowupCount = parseInt(cr.data.values?.[0]?.[0]) || 0;
    } catch (e) {}
    const newFollowupCount = currentFollowupCount + 1;
    updates.push({
      range: `'${NBD_SHEET_NAME}'!Z${rowIndex}`,
      values: [[newFollowupCount]],
    });

    // ✅ FIX: loggerStatus अलग track करो — हर case में सही value
    let loggerStatus = "Done";

    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      // ✅ Reschedule या COLD दोनों यहाँ आते हैं (rescheduleDate है)
      const isColdAction = actionType === "cold";
      const isNextFollowup = actionType === "next-followup";

      if (isColdAction) {
        loggerStatus = "COLD";
      } else if (isNextFollowup) {
        loggerStatus = "Next Followup Required";
      } else {
        loggerStatus = "Rescheduled";
      }

      let statusW = "Rescheduled";
      if (isColdAction) statusW = "COLD";
      else if (isNextFollowup) statusW = "Next Followup Required";

      updates.push({
        range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        values: [[getPlannedDateTime(rescheduleDate)]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [[statusW]],
      });
    } else if (status === "Not Interested") {
      loggerStatus = "Not Interested";

      updates.push({
        range: `'${NBD_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Not Interested"]],
      });
      if (leadInfo)
        await appendToNotInterestedSheet(
          req.sheets,
          leadInfo,
          "Step 2 - Field Visit",
          notInterestedReason || "",
          req.user?.email,
        );
    } else if (status === "Call Not Picked") {
      loggerStatus = "Call Not Picked";

      // ✅ CNP — auto +7 days to planned date
      let existingPlannedDate = "";
      try {
        const pd = await req.sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        });
        existingPlannedDate = pd.data.values?.[0]?.[0] || "";
      } catch (e) {}

      const newPlannedDate = addDaysToSheetDate(existingPlannedDate, 7);

      updates.push({
        range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        values: [[newPlannedDate]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Call Not Picked"]],
      });
    } else {
      // ✅ Done
      loggerStatus = status || "Done";

      updates.push({
        range: `'${NBD_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [[status || "Done"]],
      });

      // ✅ FSR users अपना code AN column में लिखते हैं
      if (req.user) {
        const fsrCode = getFSRCode(req.user);
        if (fsrCode) {
          updates.push({
            range: `'${NBD_SHEET_NAME}'!AN${rowIndex}`,
            values: [[fsrCode]],
          });
        }
      }
    }

    // ✅ Remarks — timestamp के साथ append
    if (remarks && String(remarks).trim() !== "") {
      const appended = await buildAppendedRemarks(
        req.sheets,
        NBD_SHEET_NAME,
        `Y${rowIndex}`,
        remarks,
      );
      if (appended)
        updates.push({
          range: `'${NBD_SHEET_NAME}'!Y${rowIndex}`,
          values: [[appended]],
        });
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    // ✅ FIX: Logger में सही status — loggerStatus use करो, "Done" नहीं
    if (leadInfo)
      await appendToLogger(
        req.sheets,
        leadInfo,
        "Step 2 - Field Visit",
        loggerStatus, // ✅ यही fix है
        remarks || "",
        req.user?.email,
      );

    // ✅ Response message
    // ✅ Response message
    let responseMessage = "Field Visit Done";
    if (rescheduleDate) {
      if (actionType === "cold") {
        responseMessage = "Marked COLD — Next follow-up after 15 days";
      } else if (actionType === "next-followup") {
        responseMessage = "Next Followup Scheduled Successfully";
      } else {
        responseMessage = "Rescheduled Successfully";
      }
    } else if (status === "Not Interested") {
      responseMessage = "Marked Not Interested";
    } else if (status === "Call Not Picked") {
      responseMessage = "Marked Call Not Picked — Planned date moved +7 days";
    }

    res.json({
      success: true,
      message: responseMessage,
      newFollowupCount,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;
