const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  // CHANNEL_PARTNER: "Channel Partener Lead FMS",
};


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

// --- READ / LIST ---

router.get("/list", async (req, res) => {
  try {
    const getData = async (sheets, sheetName) => {  // ✅ sheets parameter added
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!A8:AG`, // Extended to AG
      });

      const rows = response.data.values || [];

      return rows
        .map((row, index) => {
          const getCol = (idx) => (row[idx] ? String(row[idx]).trim() : "");

          const plannedDate = getCol(26); // AA
          const actualDate = getCol(27);  // AB
          const status = getCol(28);      // AC

          // Filter: show if planned exists and actual empty OR specific status
          if ((plannedDate && !actualDate) || status === "No conversation" || status === "Next Follow Up") {

            // ===== ALL REMARKS COLUMNS =====
            const oldRemarks = getCol(11);           // L - Initial/Oldest Remarks
            const previousRemarksDate = getCol(13);  // N - Date for T remarks
            const previousRemarks = getCol(19);      // T - Previous Remarks
            const latestOldRemarks = getCol(24);     // Y - Latest Old Remarks
            const latestOldRemarksDate = getCol(21); // V - Date for Y remarks
            const currentRemarks = getCol(32);       // AG - Current/New Remarks

            // Display priority for table: AG > Y > T > L
            let displayRemarks = currentRemarks || latestOldRemarks || previousRemarks || oldRemarks;

            return {
              rowIndex: index + 8,
              sheetName,
              uniqueId: getCol(1),
              customerName: getCol(2),
              customerContact: getCol(3),
              interestedIn: getCol(4),
              projectSelection: getCol(5),
              leadSource: getCol(6),
              leadGenNumber: getCol(7),
              leadGenName: getCol(8),
              importantNote: getCol(10),
              plannedDate,
              status: status || "Pending",
              followUpCount: getCol(31) || "0", // AF

              // ===== SEND ALL REMARKS TO FRONTEND =====
              remarks: displayRemarks,                      // For table display
              oldRemarks: oldRemarks,                       // L - Initial remarks (read-only)
              previousRemarks: previousRemarks,             // T - Previous remarks (read-only)
              previousRemarksDate: previousRemarksDate,     // N - Date for T
              latestOldRemarks: latestOldRemarks,          // Y - Latest old remarks (read-only)
              latestOldRemarksDate: latestOldRemarksDate,  // V - Date for Y
            };
          }
          return null;
        })
        .filter(Boolean);
    };

    // ✅ Pass req.sheets to getData function
    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getData(req.sheets, SHEETS.END_USER),
      getData(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    allLeads.sort((a, b) => {
      const parse = (d) => {
        if (!d) return 0;
        const p = d.split(/[\/ :]/);
        return new Date(p[2], p[1] - 1, p[0], p[3] || 0, p[4] || 0).getTime();
      };
      return parse(a.plannedDate) - parse(b.plannedDate);
    });

    res.json({ success: true, data: allLeads });

  } catch (error) {
    console.error("List API Error:", error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// --- UPDATE ---

router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks, rescheduleDate, dealMeetingDate } = req.body;

    console.log("Update Payload:", JSON.stringify(req.body, null, 2));

    if (!sheetName || !rowIndex) {
      return res.status(400).json({ success: false, error: "Missing sheetName or rowIndex" });
    }

    let currentCount = 0;
    try {
      const countRes = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!AF${rowIndex}`,
      });
      const val = countRes.data.values?.[0]?.[0];
      currentCount = parseInt(val) || 0;
    } catch (e) {
      console.warn("Could not read count:", e.message);
    }

    const newCount = currentCount + 1;
    const updates = [];
    const timestamp = getCurrentTimestamp();

    // Always increment count (AF)
    updates.push({
      range: `'${sheetName}'!AF${rowIndex}`,
      values: [[newCount]],
    });

    // ===== ALWAYS UPDATE NEW REMARKS IN COLUMN AG =====
    if (remarks && String(remarks).trim()) {
      updates.push({
        range: `'${sheetName}'!AG${rowIndex}`,  // New remarks save to AG
        values: [[String(remarks).trim()]],
      });
    }

    // ── RESCHEDULE CASE ──
    if (rescheduleDate && String(rescheduleDate).trim() !== "" && ["No conversation", "Next Follow Up"].includes(status || "")) {
      const formatted = formatDateToSheetStyle(rescheduleDate);
      console.log("Processing RESCHEDULE to:", formatted);

      updates.push({
        range: `'${sheetName}'!AA${rowIndex}`,  // Planned
        values: [[formatted]],
      });

      updates.push({
        range: `'${sheetName}'!AE${rowIndex}`,  // Next FollowUP Date
        values: [[formatted]],
      });

      updates.push({
        range: `'${sheetName}'!AC${rowIndex}`,  // Status
        values: [[status]],
      });
    }
    // ── MARK DONE / NOT INTERESTED / OTHER STATUS ──
    else {
      console.log("Processing MARK DONE / STATUS UPDATE:", status);

      // Write actual timestamp (AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[timestamp]],
      });

      // Status (AC)
      let finalStatus = status || "Done";
      updates.push({
        range: `'${sheetName}'!AC${rowIndex}`,
        values: [[finalStatus]],
      });

      if (dealMeetingDate) {
        const formattedDeal = formatDateToSheetStyle(dealMeetingDate);
        updates.push({
          range: `'${sheetName}'!AD${rowIndex}`,
          values: [[formattedDeal]],
        });
      }
    }

    if (updates.length > 0) {
      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: updates,
        },
      });
      console.log(`Updated ${updates.length} cells in row ${rowIndex}`);
    } else {
      console.log("No updates to perform");
    }

    res.json({ success: true, message: "Updated", newFollowUpCount: newCount });

  } catch (error) {
    console.error("Update Error:", error);
    res.status(500).json({ success: false, error: error.message });
  }
});

module.exports = router;