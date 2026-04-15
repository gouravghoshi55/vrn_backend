const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NBD_SHEET_NAME = "END USER LEADS FMS";
const LOGGER_SHEET_NAME = "Logger";
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

function getCurrentTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0"),
    mo = String(now.getMonth() + 1).padStart(2, "0"),
    y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0"),
    mi = String(now.getMinutes()).padStart(2, "0"),
    s = String(now.getSeconds()).padStart(2, "0");
  return `${d}/${mo}/${y} ${h}:${mi}:${s}`;
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
  return `${fd} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}:${String(now.getSeconds()).padStart(2, "0")}`;
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
    "bdm6@company.com": "BDM6",
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
  };
  return m[user.email?.toLowerCase()] || null;
}

// ✅ Read existing remarks, prepend new with timestamp
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
      const fsrDoer = row[39] ? row[39].trim() : "";

      // ✅ Only show empty status or "Rescheduled" — "Call Not Picked" automatically excluded
      const showRow = !status || status.trim().toLowerCase() === "rescheduled";

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
          fsrDoer,
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
    } = req.body;
    if (!rowIndex)
      return res
        .status(400)
        .json({ success: false, error: "Missing rowIndex" });

    const updates = [];
    const timestamp = getCurrentTimestamp();

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

    if (rescheduleDate && String(rescheduleDate).trim() !== "") {
      // ✅ Reschedule
      updates.push({
        range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        values: [[getPlannedDateTime(rescheduleDate)]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Rescheduled"]],
      });
    } else if (status === "Not Interested") {
      // ✅ Not Interested
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
      // ✅ CNP — sirf W column mein status, V (actual date) mat likho
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Call Not Picked"]],
      });
    } else {
      // ✅ Done (or any other status)
      updates.push({
        range: `'${NBD_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [[status || "Done"]],
      });
      if (req.user && req.user.assignedModule === "fsr") {
        const fc = getFSRCode(req.user);
        if (fc)
          updates.push({
            range: `'${NBD_SHEET_NAME}'!AN${rowIndex}`,
            values: [[fc]],
          });
      }
    }

    // ✅ Remarks — APPEND with timestamp
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
    if (leadInfo)
      await appendToLogger(
        req.sheets,
        leadInfo,
        "Step 2 - Field Visit",
        status || "Done",
        remarks || "",
        req.user?.email,
      );

    res.json({
      success: true,
      message: rescheduleDate
        ? "Rescheduled"
        : status === "Not Interested"
          ? "Marked Not Interested"
          : status === "Call Not Picked"
            ? "Marked Call Not Picked"
            : "Field Visit Done",
      newFollowupCount,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;
