const { getCurrentTimestamp } = require("../../utils/dateUtils");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

// GET — Show rows where Planned(H) NOT NULL & Actual(I) IS NULL
exports.getMeetingsSub = async (req, res) => {
  try {
    // ✅ Use req.sheets (retry-enabled, always defined)
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

      // ✅ Show only if Planned exists AND Actual is empty
      const hasPlanned = planned && planned.toString().trim() !== "";
      const hasActual = actual && actual.toString().trim() !== "";

      // ⚠️ NOTE: Aapke original code mein bug tha — `if (hasPlanned || hasActual) return;`
      // Iska matlab Planned hai to bhi skip kar dega. Should be:
      if (!hasPlanned || hasActual) return;

      // Decide which date to display
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

// POST — Save Meeting action
exports.submitMeetingsSubAction = async (req, res) => {
  try {
    const { rowNumber, status, channelPartnerName, reviseDate, remark } =
      req.body;

    if (!rowNumber || !status) {
      return res
        .status(400)
        .json({ success: false, message: "rowNumber and status required" });
    }

    // ✅ Use req.sheets
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

    if (status === "Done") {
      actualValue = getCurrentTimestamp();
    }

    // ✅ Use req.sheets
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
            remark || "",
          ],
        ],
      },
    });

    res.json({ success: true, message: "Meeting updated" });
  } catch (err) {
    console.error("MeetingsSub POST error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};