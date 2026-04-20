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

function getDoerTag(user) {
  if (!user) return null;
  if (user.role === "admin" || user.assignedModule === "all") return null;
  if (user.assignedModule === "fsr") return null;
  const m = {
    "bdm1@company.com": "BDM1",
    "bdm2@company.com": "BDM2",
    "bdm3@company.com": "BDM3",
    "bdm6@company.com": "BDM6",
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

// ============================================
// GET /list — Fetch all CNP leads
// ============================================
router.get("/list", async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${NBD_SHEET_NAME}'!A8:AN`,
    });
    const rows = response.data.values || [];
    const doerTag = getDoerTag(req.user);
    const leads = [];

    rows.forEach((row, index) => {
      const status = row[22] ? row[22].trim() : "";
      if (status.toLowerCase() !== "call not picked") return;

      const doer = row[38] ? row[38].trim() : "";
      if (doerTag && doer !== doerTag) return;

      leads.push({
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
        plannedDate: row[20] ? row[20].trim() : "",
        status,
        followupCount: row[25] ? parseInt(row[25].trim(), 10) || 0 : 0,
        remarks:
          (row[24] ? row[24].trim() : "") ||
          (row[19] ? row[19].trim() : "") ||
          (row[11] ? row[11].trim() : ""),
        oldRemarks: row[11] ? row[11].trim() : "",
        previousRemarks: row[19] ? row[19].trim() : "",
        doer,
      });
    });

    leads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    res
      .status(500)
      .json({ success: false, error: "Failed", message: error.message });
  }
});

// ============================================
// POST /update — Handle CNP actions
// ============================================
router.post("/update", async (req, res) => {
  try {
    const {
      rowIndex,
      action,
      fieldVisitDate,
      nextFollowUpDate,
      remarks,
      notInterestedReason,
      leadInfo,
    } = req.body;

    if (!rowIndex || !action) {
      return res
        .status(400)
        .json({ success: false, error: "Missing rowIndex or action" });
    }

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // Increment followup count
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

    let logStep = "CNP";
    let logStatus = "";
    let responseMessage = "";

    if (action === "schedule") {
      // Schedule Site Visit — lead goes back to Field Visit
      if (!fieldVisitDate) {
        return res
          .status(400)
          .json({
            success: false,
            error: "fieldVisitDate required for schedule action",
          });
      }
      updates.push({
        range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        values: [[getPlannedDateTime(fieldVisitDate)]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Rescheduled"]],
      });
      logStep = "CNP - Schedule Site Visit";
      logStatus = "Rescheduled";
      responseMessage = "Site Visit scheduled — lead moved to Field Visit";
    } else if (action === "not-interested") {
      // Not Interested — close lead
      if (!notInterestedReason || !String(notInterestedReason).trim()) {
        return res
          .status(400)
          .json({
            success: false,
            error: "Reason required for Not Interested",
          });
      }
      updates.push({
        range: `'${NBD_SHEET_NAME}'!V${rowIndex}`,
        values: [[timestamp]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Not Interested"]],
      });
      if (leadInfo) {
        await appendToNotInterestedSheet(
          req.sheets,
          leadInfo,
          "CNP - Not Interested",
          notInterestedReason,
          req.user?.email,
        );
      }
      logStep = "CNP - Not Interested";
      logStatus = "Not Interested";
      responseMessage = "Lead marked Not Interested";
    } else if (action === "cnp-again") {
      // CNP again — stay in CNP with new planned date
      if (!nextFollowUpDate) {
        return res
          .status(400)
          .json({
            success: false,
            error: "nextFollowUpDate required for cnp-again action",
          });
      }
      updates.push({
        range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        values: [[getPlannedDateTime(nextFollowUpDate)]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Call Not Picked"]],
      });
      logStep = "CNP - Call Not Picked";
      logStatus = "Call Not Picked";
      responseMessage = "Lead updated — next follow-up scheduled";
    } else if (action === "next-followup") {
      // Next Follow Up — stay in CNP with new planned date (Option B)
      if (!nextFollowUpDate) {
        return res
          .status(400)
          .json({
            success: false,
            error: "nextFollowUpDate required for next-followup action",
          });
      }
      updates.push({
        range: `'${NBD_SHEET_NAME}'!U${rowIndex}`,
        values: [[getPlannedDateTime(nextFollowUpDate)]],
      });
      updates.push({
        range: `'${NBD_SHEET_NAME}'!W${rowIndex}`,
        values: [["Call Not Picked"]],
      });
      logStep = "CNP - Next Follow Up";
      logStatus = "Call Not Picked";
      responseMessage = "Next follow-up scheduled — lead stays in CNP";
    } else {
      return res
        .status(400)
        .json({
          success: false,
          error:
            "Invalid action. Use: schedule, not-interested, cnp-again, next-followup",
        });
    }

    // Append remarks
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

    // Batch update
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    // Logger
    if (leadInfo) {
      await appendToLogger(
        req.sheets,
        leadInfo,
        logStep,
        logStatus,
        remarks || "",
        req.user?.email,
      );
    }

    res.json({ success: true, message: responseMessage, newFollowupCount });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;
