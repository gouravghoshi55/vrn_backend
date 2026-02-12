const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// ============================================
// Helpers
// ============================================

function getCurrentTimestamp() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0");
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const y = now.getFullYear();
  const h = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const s = String(now.getSeconds()).padStart(2, "0");
  return `${d}/${m}/${y} ${h}:${mi}:${s}`;
}

function getPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  const now = new Date();
  const timePart = `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}:${String(now.getSeconds()).padStart(2, "0")}`;

  let formattedDate = dateStr;
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts[0].length === 4) {
      formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  return `${formattedDate} ${timePart}`;
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

// ============================================
// FETCH LIST (GET)
// ============================================

async function getFilteredLeads(sheets, sheetName) {
  // Range badha kar AE (Index 30) tak kar di hai
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A8:AE`,
  });

  const rows = response.data.values || [];
  const filtered = [];

  rows.forEach((row, index) => {
    // Indices based on Screenshot:
    // X=23 (Field Visit Remarks), Y=24 (Planned), Z=25 (Actual), AA=26 (Status)
    // AD=29 (Count), AE=30 (Current Step Remarks)

    const plannedDate = row[24] ? row[24].trim() : ""; // Column Y
    const actualDate = row[25] ? row[25].trim() : "";  // Column Z
    const status = row[26] ? row[26].trim() : "";      // Column AA
    const followUpCount = row[29] ? row[29].trim() : "0"; // Column AD
    
    // CHANGE: Hum 'X' (Index 23) padhenge taaki Sales person pichla remark dekh sake
    const previousRemarks = row[23] ? row[23].trim() : ""; 

    // Filter Logic: Show if Planned exists AND Actual is empty
    if (plannedDate && !actualDate) {
      filtered.push({
        rowIndex: index + 8,
        sheetName,
        uniqueId: row[1] || "",
        customerName: row[2] || "",
        customerContact: row[3] || "",
        projectSelection: row[5] || "",
        plannedDate,
        status: status || "Pending",
        followUpCount: parseInt(followUpCount) || 0,
        remarks: previousRemarks, // Frontend ko hum Column X bhej rahe hain
      });
    }
  });

  return filtered;
}

// ============================================
// ROUTES
// ============================================

router.get("/list", async (req, res) => {
  try {
    const [endUser, channelPartner] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);
    const all = [...endUser, ...channelPartner];
    all.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));
    
    res.json({ success: true, data: all, total: all.length });
  } catch (err) {
    console.error("Error fetching data:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ============================================
// UPDATE (POST)
// ============================================

router.post("/update", async (req, res) => {
  try {
    const { sheetName, rowIndex, status, dealMeetingDate, nextFollowUpDate, remarks, currentFollowUpCount } = req.body;

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({ success: false, error: "Missing fields" });
    }

    const normalizedStatus = status.trim().toLowerCase();
    const timestamp = getCurrentTimestamp();
    const updates = [];

    // --- CASE 1: NO CONVERSATION (Reschedule) ---
    if (normalizedStatus === "no conversation") {
      const newCount = (parseInt(currentFollowUpCount) || 0) + 1;
      
      // Update Planned (Column Y)
      if (nextFollowUpDate) {
        updates.push({ range: `'${sheetName}'!Y${rowIndex}`, values: [[getPlannedDateTime(nextFollowUpDate)]] });
      }
      // Update Actual (Column Z) -> Loop back ke liye Actual clear rakhna hota hai usually, 
      // par agar aap history maintain kar rahe ho toh timestamp daal sakte ho. 
      // Screenshot pattern ke hisab se ye logic aapke flow chart par depend karta hai.
      // Assuming Loop back logic: We typically DON'T fill Actual if we want it to reappear, 
      // BUT your code fills Actual. I will follow your code logic but fix Columns.

      updates.push({ range: `'${sheetName}'!Z${rowIndex}`, values: [[timestamp]] }); // Actual
      updates.push({ range: `'${sheetName}'!AA${rowIndex}`, values: [["No conversation"]] }); // Status
      
      if (nextFollowUpDate) {
        updates.push({ range: `'${sheetName}'!AC${rowIndex}`, values: [[nextFollowUpDate]] }); // Next FollowUp Date (Col AC)
      }
      updates.push({ range: `'${sheetName}'!AD${rowIndex}`, values: [[newCount.toString()]] }); // Count (Col AD)
      
      if (remarks) {
        updates.push({ range: `'${sheetName}'!AE${rowIndex}`, values: [[remarks]] }); // Remarks (Col AE)
      }
    }

    // --- CASE 2: DONE ---
    else if (normalizedStatus === "done") {
      updates.push({ range: `'${sheetName}'!Z${rowIndex}`, values: [[timestamp]] }); // Actual
      updates.push({ range: `'${sheetName}'!AA${rowIndex}`, values: [["Done"]] });   // Status
      
      if (dealMeetingDate) {
        updates.push({ range: `'${sheetName}'!AB${rowIndex}`, values: [[dealMeetingDate]] }); // Deal Date (Col AB)
      }
      if (remarks) {
        updates.push({ range: `'${sheetName}'!AE${rowIndex}`, values: [[remarks]] }); // Remarks (Col AE)
      }
    }

    // --- CASE 3: OTHERS ---
    else {
      updates.push({ range: `'${sheetName}'!Z${rowIndex}`, values: [[timestamp]] }); // Actual
      updates.push({ range: `'${sheetName}'!AA${rowIndex}`, values: [[status]] });   // Status
      
      if (remarks) {
        updates.push({ range: `'${sheetName}'!AE${rowIndex}`, values: [[remarks]] }); // Remarks (Col AE)
      }
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "USER_ENTERED", data: updates },
    });

    res.json({ success: true, message: "Follow Up Updated Successfully" });
  } catch (err) {
    console.error("❌ Error updating:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;