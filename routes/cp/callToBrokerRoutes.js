const express = require("express");
const router = express.Router();

const CP_OUTGOING_SPREADSHEET_ID = process.env.CP_OUTGOING_SPREADSHEET_ID;
const SHEET_NAME = "FMS";

// --- HELPER FUNCTIONS ---

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

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const p = dateStr.split(/[\/\-]/);
  if (p.length === 3) {
    if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]);
    return new Date(p[2], p[1] - 1, p[0]);
  }
  return new Date(dateStr);
}

function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  if (dateStr.includes("T")) {
    const [dt, t] = dateStr.split("T");
    const [y, m, d] = dt.split("-");
    return `${d}/${m}/${y} ${t}:00`;
  }
  const now = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
  let fd = dateStr;
  if (dateStr.includes("-")) {
    const p = dateStr.split("-");
    if (p[0].length === 4) fd = `${p[2]}/${p[1]}/${p[0]}`;
  }
  return `${fd} ${String(now.getUTCHours()).padStart(2, "0")}:${String(now.getUTCMinutes()).padStart(2, "0")}:${String(now.getUTCSeconds()).padStart(2, "0")}`;
}

// ✅ Append-only remarks with timestamp
async function buildAppendedRemarks(sheets, cellRange, newRemark) {
  if (!newRemark || !String(newRemark).trim()) return null;
  let existing = "";
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: CP_OUTGOING_SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!${cellRange}`,
    });
    existing = r.data.values?.[0]?.[0] || "";
  } catch (e) {}
  const timestamped = `${getShortTimestamp()} ${String(newRemark).trim()}`;
  return existing.trim() ? `${existing.trim()}\n${timestamped}` : timestamped;
}

// ============================================
// GET /list — Fetch Call to Broker leads
// ============================================
router.get("/list", async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: CP_OUTGOING_SPREADSHEET_ID,
      range: `'${SHEET_NAME}'!A8:P`,
    });
    const rows = response.data.values || [];
    const leads = [];

    rows.forEach((row, index) => {
      const g = (i) => (row[i] ? String(row[i]).trim() : "");
      const plannedDate = g(6);
      const actualDate = g(7);
      const status = g(8);

      const showRow =
        (plannedDate && !actualDate) ||
        status === "Call Again" ||
        status === "No Connection";

      if (!showRow) return;

      leads.push({
        rowIndex: index + 8,
        uniqueId: g(1),
        firmName: g(2),
        contact: g(3),
        locality: g(4),
        callerName: g(5),
        plannedDate,
        actualDate,
        status: status || "Pending",
        nextFollowUpDate: g(9),
        isLeadQualified: g(10),
        contactPersonName: g(11),
        rera: g(12),
        remark: g(13),
        followUpCounter: parseInt(g(14)) || 0,
        meetingDate: g(15),
      });
    });

    // Sort by planned date (earliest first)
    // In backend /list — at sort line
    leads.sort(
      (a, b) =>
        parseDate(a.nextFollowUpDate || a.plannedDate) -
        parseDate(b.nextFollowUpDate || b.plannedDate),
    );

    res.json({ success: true, data: leads, total: leads.length });
  } catch (error) {
    console.error("❌ Error fetching Call to Broker:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to fetch data",
      message: error.message,
    });
  }
});

// ============================================
// POST /update — Update Call to Broker lead
// ============================================
router.post("/update", async (req, res) => {
  try {
    const {
      rowIndex,
      status,
      nextFollowUpDate,
      isLeadQualified,
      contactPersonName,
      rera,
      remarks,
      meetingDate,
    } = req.body;

    if (!rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing rowIndex or status",
      });
    }

    const updates = [];
    const timestamp = getCurrentTimestamp();

    // ✅ Status (Col I, index 8)
    updates.push({
      range: `'${SHEET_NAME}'!I${rowIndex}`,
      values: [[status]],
    });

    // ✅ Actual Date logic — ALWAYS runs regardless of nextFollowUpDate
    // ✅ Actual Date logic — ALWAYS runs regardless of nextFollowUpDate
    if (status === "Call Again" || status === "No Connection") {
      // Call Again / No Connection: clear actual so lead shows again with new planned date
      updates.push({
        range: `'${SHEET_NAME}'!H${rowIndex}`,
        values: [[""]],
      });
    } else {
      // All other statuses: write actual timestamp
      updates.push({
        range: `'${SHEET_NAME}'!H${rowIndex}`,
        values: [[timestamp]],
      });
    }

    // ✅ NEW: No Connection — auto +15 days planned date
    if (status === "No Connection") {
      // Generate date 15 days from now at 10:00 AM IST
      const futureDate = new Date(new Date().getTime() + 5.5 * 60 * 60 * 1000);
      futureDate.setUTCDate(futureDate.getUTCDate() + 15);

      const dd = String(futureDate.getUTCDate()).padStart(2, "0");
      const mm = String(futureDate.getUTCMonth() + 1).padStart(2, "0");
      const yyyy = futureDate.getUTCFullYear();
      const noConnectionDate = `${dd}/${mm}/${yyyy} 10:00:00`;

      // Write to BOTH G (Planned) and J (Next Follow Up) for consistency
      updates.push({
        range: `'${SHEET_NAME}'!G${rowIndex}`,
        values: [[noConnectionDate]],
      });
      updates.push({
        range: `'${SHEET_NAME}'!J${rowIndex}`,
        values: [[noConnectionDate]],
      });
    }

    // ✅ Next Follow Up Date (Col J, index 9)
    if (nextFollowUpDate) {
      updates.push({
        range: `'${SHEET_NAME}'!J${rowIndex}`,
        values: [[getPlannedDateTime(nextFollowUpDate)]],
      });
      // ✅ Also update Planned Date (Col G) for next cycle
      updates.push({
        range: `'${SHEET_NAME}'!G${rowIndex}`,
        values: [[getPlannedDateTime(nextFollowUpDate)]],
      });
    }

    // ✅ Is Lead Qualified (Col K, index 10)
    if (isLeadQualified !== undefined && isLeadQualified !== null) {
      updates.push({
        range: `'${SHEET_NAME}'!K${rowIndex}`,
        values: [[isLeadQualified]],
      });
    }

    // ✅ Contact Person Name (Col L, index 11)
    if (contactPersonName !== undefined) {
      updates.push({
        range: `'${SHEET_NAME}'!L${rowIndex}`,
        values: [[contactPersonName]],
      });
    }

    // ✅ RERA (Col M, index 12)
    if (rera !== undefined && rera !== null) {
      updates.push({
        range: `'${SHEET_NAME}'!M${rowIndex}`,
        values: [[rera]],
      });
    }

    // ✅ Remarks — APPEND with timestamp (Col N, index 13)
    if (remarks && String(remarks).trim() !== "") {
      const appended = await buildAppendedRemarks(
        req.sheets,
        `N${rowIndex}`,
        remarks,
      );
      if (appended) {
        updates.push({
          range: `'${SHEET_NAME}'!N${rowIndex}`,
          values: [[appended]],
        });
      }
    }

    // ✅ Follow Up Counter (Col O, index 14) — +1 only when "Call Again"
    if (status === "Call Again" || status === "No Connection") {
      let currentCount = 0;
      try {
        const cr = await req.sheets.spreadsheets.values.get({
          spreadsheetId: CP_OUTGOING_SPREADSHEET_ID,
          range: `'${SHEET_NAME}'!O${rowIndex}`,
        });
        currentCount = parseInt(cr.data.values?.[0]?.[0]) || 0;
      } catch (e) {}
      updates.push({
        range: `'${SHEET_NAME}'!O${rowIndex}`,
        values: [[currentCount + 1]],
      });
    }

    // ✅ Meeting Date (Col P, index 15) — only for "Agreed to next meeting"
    if (status === "Agreed to next meeting" && meetingDate) {
      updates.push({
        range: `'${SHEET_NAME}'!P${rowIndex}`,
        values: [[getPlannedDateTime(meetingDate)]],
      });
    }


    // Batch update
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: CP_OUTGOING_SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    res.json({
      success: true,
      message:
        status === "Call Again"
          ? "Follow-up scheduled"
          : status === "No Connection"
            ? "Marked No Connection — Next call after 15 days"
            : status === "Agreed to next meeting"
              ? "Meeting scheduled"
              : `Marked as ${status}`,
    });
  } catch (error) {
    console.error("❌ Error updating Call to Broker:", error.message);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

module.exports = router;
