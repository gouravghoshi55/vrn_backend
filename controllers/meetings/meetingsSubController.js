const { sheets } = require("../../config/googleAuth");
const { getCurrentTimestamp } = require("../../utils/dateUtils");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

// GET — Show rows where Planned(O) NOT NULL & Actual(P) IS NULL
exports.getMeetingsSub = async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:V`,
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, idx) => {
      const planned     = row[14]; // O
      const actual      = row[15]; // P
      const status      = (row[16] || "").toString().trim(); // Q
      const reviseDate  = row[18]; // S
      const reviseCount = row[19]; // T

      // ✅ Show only if Planned exists AND Actual is empty
      const hasPlanned = planned && planned.toString().trim() !== "";
      const hasActual  = actual && actual.toString().trim() !== "";

      if (!hasPlanned || hasActual) return;

      // ✅ Decide which date to display
      // If status is "Revise" and reviseDate exists → use reviseDate
      // Else → use planned
      const displayDate = (status.toLowerCase() === "revise" && reviseDate)
        ? reviseDate
        : planned;

      filtered.push({
        rowNumber: DATA_START_ROW + idx,
        uniqueId: row[1] || "",
        firmName: row[2] || "",
        contact:  row[3] || "",
        locality: row[4] || "",
        plannedDate: displayDate,        // current applicable date
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

// POST — Save Meeting action
exports.submitMeetingsSubAction = async (req, res) => {
  try {
    const { rowNumber, status, channelPartnerName, reviseDate, remark } = req.body;

    if (!rowNumber || !status) {
      return res.status(400).json({ success: false, message: "rowNumber and status required" });
    }

    // Read existing P:U row
    const currentRow = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!P${rowNumber}:U${rowNumber}`,
    });
    const existing = (currentRow.data.values && currentRow.data.values[0]) || [];
    const currentReviseCount = parseInt(existing[4] || "0", 10) || 0;

    let newReviseCount = currentReviseCount;
    let newReviseDate  = existing[3] || "";
    let actualValue    = existing[0] || "";

    if (status === "Revise") {
      newReviseCount = currentReviseCount + 1;
      newReviseDate = reviseDate || "";
    }

    if (status === "Done") {
      actualValue = getCurrentTimestamp();
    }

    // Update P (Actual), Q (Status), R (CP Name), S (Revise Date), T (Revise Count), U (Remark)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!P${rowNumber}:U${rowNumber}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          actualValue,
          status,
          channelPartnerName || "",
          newReviseDate,
          newReviseCount,
          remark || "",
        ]],
      },
    });

    res.json({ success: true, message: "Meeting updated" });
  } catch (err) {
    console.error("MeetingsSub POST error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};