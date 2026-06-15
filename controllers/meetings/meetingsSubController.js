const { getCurrentTimestamp } = require("../../utils/dateUtils");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

// ✅ NEW — Not Interested logging goes to main working sheet
const MAIN_SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

// ============================================
// GET — Show rows where Planned(H) NOT NULL & Actual(I) IS NULL
// ============================================
exports.getMeetingsSub = async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:N`,
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, idx) => {
      const planned = row[7]; // H
      const actual = row[8]; // I
      const status = (row[9] || "").toString().trim(); // J
      const reviseDate = row[11]; // L
      const reviseCount = row[12]; // M

      const hasPlanned = planned && planned.toString().trim() !== "";
      const hasActual = actual && actual.toString().trim() !== "";

      if (!hasPlanned || hasActual) return;

      const displayDate =
        status.toLowerCase() === "revise" && reviseDate ? reviseDate : planned;

      filtered.push({
        rowNumber: DATA_START_ROW + idx,
        uniqueId: row[1] || "",
        firmName: row[2] || "",
        contact: row[3] || "",
        locality: row[4] || "",
        plannedDate: displayDate,
        status: status || "",
        reviseCount: reviseCount || "0",
      });
    });

    res.json({ success: true, data: filtered });
  } catch (err) {
    console.error("MeetingsSub GET error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};

// ============================================
// POST — Save Meeting action (with Not Interested support)
// ============================================
exports.submitMeetingsSubAction = async (req, res) => {
  try {
    const {
      rowNumber,
      status,
      channelPartnerName,
      reviseDate,
      remark,
      notInterestedReason, // ✅ NEW
    } = req.body;

    if (!rowNumber || !status) {
      return res
        .status(400)
        .json({ success: false, message: "rowNumber and status required" });
    }

    // ✅ Validate Not Interested has reason
    if (status === "Not Interested" && !notInterestedReason?.trim()) {
      return res.status(400).json({
        success: false,
        message: "Reason required for Not Interested status",
      });
    }

    // ✅ Read existing I:N row
    const currentRow = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!I${rowNumber}:N${rowNumber}`,
    });
    const existing =
      (currentRow.data.values && currentRow.data.values[0]) || [];
    const currentReviseCount = parseInt(existing[4] || "0", 10) || 0;

    let newReviseCount = currentReviseCount;
    let newReviseDate = existing[3] || "";
    let actualValue = existing[0] || "";

    if (status === "Revise") {
      newReviseCount = currentReviseCount + 1;
      newReviseDate = reviseDate || "";
    }

    // ✅ Done, Not Done, Not Interested — all write timestamp to Actual (Col I)
    if (
      status === "Done" ||
      status === "Not Done" ||
      status === "Not Interested"
    ) {
      actualValue = getCurrentTimestamp();
    }

    // ✅ Build final remark — for Not Interested, append reason
    const finalRemark =
      status === "Not Interested"
        ? `${remark || ""}${remark ? " | " : ""}Reason: ${notInterestedReason}`
        : remark || "";

    // ✅ Update I (Actual), J (Status), K (CP Name), L (Revise Date), M (Revise Count), N (Remark)
    await req.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!I${rowNumber}:N${rowNumber}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [
          [
            actualValue,
            status,
            channelPartnerName || "",
            newReviseDate,
            newReviseCount,
            finalRemark,
          ],
        ],
      },
    });

    // ✅ NEW: Log Not Interested to dedicated sheet (best-effort, non-blocking)
    if (status === "Not Interested") {
      try {
        // Read lead info from columns A-E for context
        const leadInfoRow = await req.sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `${SHEET_NAME}!A${rowNumber}:E${rowNumber}`,
        });
        const leadCols = leadInfoRow.data.values?.[0] || [];
        const uniqueId = leadCols[1] || "";
        const firmName = leadCols[2] || "";
        const contact = leadCols[3] || "";
        const locality = leadCols[4] || "";

        const ts = getCurrentTimestamp();
        await req.sheets.spreadsheets.values.append({
          spreadsheetId: MAIN_SPREADSHEET_ID,
          range: `'${NOT_INTERESTED_SHEET}'!A:K`,
          valueInputOption: "USER_ENTERED",
          insertDataOption: "INSERT_ROWS",
          requestBody: {
            values: [
              [
                ts,
                "Meetings Sub - Not Interested",
                uniqueId,
                firmName,
                contact,
                locality,
                "", // project (N/A here)
                "", // leadSource (N/A here)
                channelPartnerName || "",
                notInterestedReason,
                req.user?.email || "",
              ],
            ],
          },
        });
        console.log(`✅ Not Interested logged: ${uniqueId} - ${firmName}`);
      } catch (logErr) {
        console.warn("⚠️ NI log failed:", logErr.message);
        // Don't fail main request if logging fails
      }
    }

    // ✅ Custom success message
    let message = "Meeting updated";
    if (status === "Not Interested") {
      message = "Marked Not Interested — logged to NI sheet";
    } else if (status === "Done") {
      message = "Meeting marked Done";
    } else if (status === "Revise") {
      message = "Meeting revised — new date scheduled";
    } else if (status === "Not Done") {
      message = "Meeting marked Not Done";
    }

    res.json({ success: true, message });
  } catch (err) {
    console.error("MeetingsSub POST error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};