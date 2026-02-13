const express = require("express");
const router = express.Router();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// --- HELPER FUNCTIONS ---

// Current Timestamp (DD/MM/YYYY HH:mm:ss)
function getCurrentTimestamp() {
  const now = new Date();
  return formatDateToSheetStyle(now);
}

// Universal Date Formatter
// Input: Date Object, ISO String, or YYYY-MM-DD string
// Output: DD/MM/YYYY HH:mm:ss
function formatDateToSheetStyle(dateInput) {
  if (!dateInput) return "";

  const d = new Date(dateInput);

  // Check for Invalid Date
  if (isNaN(d.getTime())) return "";

  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0"); // Month is 0-indexed
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");
  const seconds = String(d.getSeconds()).padStart(2, "0");

  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

// --- READ DATA ROUTE ---
router.get("/list", async (req, res) => {
  try {
    const getData = async (sheetName) => {
      const response = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!A8:AF`, // Fetch up to Remarks
      });

      const rows = response.data.values || [];
      return rows
        .map((row, index) => {
          // Mapping based on your structure
          const plannedDate = row[25] ? row[25].trim() : ""; // Z
          const actualDate = row[26] ? row[26].trim() : "";  // AA
          const status = row[27] ? row[27].trim() : "";      // AB

          // Filter Logic
          if ((plannedDate && !actualDate) || status === "No conversation") {
            return {
              rowIndex: index + 8,
              sheetName: sheetName,
              uniqueId: row[1] || "",
              customerName: row[2] || "",
              customerContact: row[3] || "",
              interestedIn: row[4] || "",
              projectSelection: row[5] || "",
              leadSource: row[6] || "",
              plannedDate: plannedDate,
              status: status || "Pending",
              followUpCount: row[30] || "0", // AE
              remarks: row[31] || "",        // AF
            };
          }
          return null;
        })
        .filter(item => item !== null); // Remove nulls
    };

    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getData(SHEETS.END_USER),
      getData(SHEETS.CHANNEL_PARTNER),
    ]);

    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    // Sort logic safe check
    allLeads.sort((a, b) => {
      // Custom date parser for DD/MM/YYYY to ensure sorting works
      const parse = (d) => {
        if (!d) return 0;
        const p = d.split(/[\/\- :]/);
        return new Date(p[2], p[1] - 1, p[0], p[3] || 0, p[4] || 0).getTime();
      };
      return parse(a.plannedDate) - parse(b.plannedDate);
    });

    res.json({ success: true, data: allLeads });

  } catch (error) {
    console.error("❌ List API Error:", error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// --- UPDATE ROUTE (ROBUST VERSION) ---
router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, remarks, rescheduleDate, dealMeetingDate } = req.body;

    console.log("📝 Incoming Update Payload:", JSON.stringify(req.body, null, 2));

    if (!sheetName || !rowIndex) {
      throw new Error("Missing sheetName or rowIndex");
    }

    // 1. Safe Fetch for FollowUP Count
    let currentCount = 0;
    try {
      const countResponse = await req.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${sheetName}'!AE${rowIndex}`,
      });
      if (countResponse.data.values && countResponse.data.values[0]) {
        const val = countResponse.data.values[0][0];
        currentCount = isNaN(parseInt(val)) ? 0 : parseInt(val);
      }
    } catch (err) {
      console.warn("⚠️ Could not fetch count, starting from 0", err.message);
    }

    const newCount = currentCount + 1;
    const updates = [];

    // --- A. ALWAYS UPDATE THESE ---

    // Update Count (AE)
    updates.push({
      range: `'${sheetName}'!AE${rowIndex}`,
      values: [[newCount]]
    });

    // Update Remarks (AF)
    if (remarks !== undefined) {
      updates.push({
        range: `'${sheetName}'!AF${rowIndex}`,
        values: [[remarks]]
      });
    }

    // --- B. CONDITIONAL LOGIC ---

    if (rescheduleDate) {
      // === RESCHEDULE SCENARIO ===
      const formattedDate = formatDateToSheetStyle(rescheduleDate);
      console.log(`📅 Processing Reschedule: ${rescheduleDate} -> ${formattedDate}`);

      // Update Next FollowUp Date (AD)
      updates.push({
        range: `'${sheetName}'!AD${rowIndex}`,
        values: [[formattedDate]],
      });

      // Override Planned Date (Z)
      updates.push({
        range: `'${sheetName}'!Z${rowIndex}`,
        values: [[formattedDate]],
      });

      // Update Status (AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[status]],
      });

    } else {
      // === DONE / STATUS UPDATE SCENARIO ===
      const timestamp = getCurrentTimestamp();

      // Update Actual Time (AA)
      updates.push({
        range: `'${sheetName}'!AA${rowIndex}`,
        values: [[timestamp]],
      });

      // Update Status (AB)
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[status]],
      });

      // Update Deal Meeting Date (AC) - Optional check
      if (dealMeetingDate) {
        const formattedDealDate = formatDateToSheetStyle(dealMeetingDate);
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[formattedDealDate]],
        });
      }
    }

    // --- C. EXECUTE BATCH UPDATE ---
    if (updates.length > 0) {
      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: updates
        },
      });
    }

    console.log(`✅ Success: Updated Row ${rowIndex} in ${sheetName}`);
    res.json({ success: true, message: "Update successful" });

  } catch (error) {
    console.error("❌ Update API Crash:", error); // Ye error server terminal me dekhein
    res.status(500).json({
      success: false,
      error: "Internal Server Error",
      message: error.message
    });
  }
});

module.exports = router;