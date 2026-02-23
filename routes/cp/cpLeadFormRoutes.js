const express = require("express");
const router = express.Router();

const CP_SPREADSHEET_ID = process.env.CP_SPREADSHEET_ID;
const CP_LEAD_SHEET = "Channel Partner FMS";
const DATA_START_ROW = 8; // Data starts from row 8

// ============================================
// Helpers
// ============================================

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

function formatDateTimeForSheet(input) {
  // Input: datetime-local string like "2026-02-25T14:30"
  if (!input || !input.includes("T")) return "";

  const [datePart, timePart] = input.split("T");
  const [year, month, day] = datePart.split("-");
  
  // Google Sheets friendly format: DD/MM/YYYY HH:mm:ss
  return `${day}/${month}/${year} ${timePart}:00`;
}

function getCurrentISTTimestamp() {
  return new Date()
    .toLocaleString("en-IN", {
      timeZone: "Asia/Kolkata",
      hour12: false,
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    })
    .replace(/(\d+)\/(\d+)\/(\d+),/, "$1/$2/$3");
}

// ============================================
// GET /list - Fetch pending leads
// ============================================

router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Lead Form pending records...");

    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: CP_SPREADSHEET_ID,
      range: `'${CP_LEAD_SHEET}'!A${DATA_START_ROW}:V`, // A to V (up to Can Contact)
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      // 0-based column indices
      const uniqueId           = row[1]  || ""; // B
      const customerName       = row[2]  || ""; // C
      const customerContact    = row[3]  || ""; // D
      const interestedIn       = row[4]  || ""; // E
      const leadGenBy          = row[5]  || ""; // F
      const leadGenNumber      = row[6]  || ""; // G
      const leadGenName        = row[7]  || ""; // H
      const leadRemark         = row[8]  || ""; // I
      const followUpCountStr   = row[9]  || "0"; // J
      const planned            = row[10] || ""; // K
      const actual             = row[11] || ""; // L
      const status             = row[12] || ""; // M
      const projectSelection   = row[13] || ""; // N
      const importantNote      = row[14] || ""; // O
      const purpose            = row[15] || ""; // P
      const nextFollowUp       = row[16] || ""; // Q
      const sendWhatsapp       = row[17] || "No"; // R
      const plannedSiteVisit   = row[18] || ""; // S
      const remark             = row[19] || ""; // T
      const notQualifiedReason = row[20] || ""; // U
      const canContact         = row[21] || ""; // V

      // Show only pending: Planned has value AND Actual is empty
      if (planned.trim() && !actual.trim()) {
        filteredLeads.push({
          rowIndex: index + DATA_START_ROW,
          uniqueId,
          customerName,
          contactNumber: customerContact,
          interestedIn,
          leadGeneratedBy: leadGenBy,
          leadGenNumber,
          leadGenName,
          leadRemark,
          followUpCount: parseInt(followUpCountStr) || 0,
          plannedDate: planned,
          status,
          projectSelection,
          importantNote,
          purpose,
          nextFollowUp,
          sendWhatsapp,
          plannedSiteVisit,
          remarks: remark,
          notQualifiedReason,
          canContact,
        });
      }
    });

    // Sort by planned date ascending
    filteredLeads.sort((a, b) => parseDate(a.plannedDate) - parseDate(b.plannedDate));

    console.log(`✅ Found ${filteredLeads.length} pending leads`);

    res.json({
      success: true,
      data: filteredLeads,
      total: filteredLeads.length,
    });
  } catch (error) {
    console.error("❌ Error fetching CP leads:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

// ============================================
// POST /update - Update lead based on status
// ============================================

router.post("/update", async (req, res) => {
  try {
    const {
      rowIndex,
      status,
      remarks = "",
      projectSelection = "",
      importantNote = "",
      purpose = "",
      plannedSiteVisit = "",        // datetime-local
      sendWhatsapp = "No",
      whatsappProject = "",
      alternateWhatsapp = "",
      nextFollowUp = "",            // datetime-local
      notQualifiedReason = "",
      canContact = "Yes",
      currentFollowUpCount = 0,
    } = req.body;

    if (!rowIndex || rowIndex < DATA_START_ROW || !status) {
      return res.status(400).json({
        success: false,
        error: "Invalid or missing rowIndex / status",
      });
    }

    console.log(`📝 Updating lead row ${rowIndex} → Status: ${status}`);

    const timestamp = getCurrentISTTimestamp();
    const newFollowUpCount = Number(currentFollowUpCount) + 1;
    const updates = [];

    // Always update these
    updates.push(
      // J - FollowUp Count
      { range: `'${CP_LEAD_SHEET}'!J${rowIndex}`, values: [[newFollowUpCount]] },
      // L - Actual (timestamp of this update)
      { range: `'${CP_LEAD_SHEET}'!L${rowIndex}`, values: [[timestamp]] },
      // M - Status
      { range: `'${CP_LEAD_SHEET}'!M${rowIndex}`, values: [[status]] },
      // R - Send Details on WhatsApp
      { range: `'${CP_LEAD_SHEET}'!R${rowIndex}`, values: [[sendWhatsapp]] }
    );

    // Remarks (T) - always update if provided
    if (remarks.trim()) {
      updates.push({
        range: `'${CP_LEAD_SHEET}'!T${rowIndex}`,
        values: [[remarks]],
      });
    }

    let whatsappNote = "";
    if (sendWhatsapp === "Yes" && whatsappProject) {
      const altNum = alternateWhatsapp.trim() || "Same as main number";
      whatsappNote = `\nWhatsApp Sent: ${whatsappProject} (Alt: ${altNum})`;
    }

    // Status-specific logic
    switch (status) {
      case "Qualified":
        if (projectSelection) {
          updates.push({ range: `'${CP_LEAD_SHEET}'!N${rowIndex}`, values: [[projectSelection]] });
        }
        if (importantNote) {
          updates.push({ range: `'${CP_LEAD_SHEET}'!O${rowIndex}`, values: [[importantNote]] });
        }
        if (purpose) {
          updates.push({ range: `'${CP_LEAD_SHEET}'!P${rowIndex}`, values: [[purpose]] });
        }
        if (plannedSiteVisit) {
          const formatted = formatDateTimeForSheet(plannedSiteVisit);
          updates.push({ range: `'${CP_LEAD_SHEET}'!S${rowIndex}`, values: [[formatted]] });
          // Also update Planned (K) for visibility
          updates.push({ range: `'${CP_LEAD_SHEET}'!K${rowIndex}`, values: [[formatted]] });
        }
        if (canContact) {
          updates.push({ range: `'${CP_LEAD_SHEET}'!V${rowIndex}`, values: [[canContact]] });
        }
        if (whatsappNote) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!T${rowIndex}`,
            values: [[(remarks + whatsappNote).trim()]],
          });
        }
        break;

      case "Next Followup Required":
        if (nextFollowUp) {
          const formatted = formatDateTimeForSheet(nextFollowUp);
          updates.push({ range: `'${CP_LEAD_SHEET}'!Q${rowIndex}`, values: [[formatted]] });
          // Update Planned (K) to next follow-up date
          updates.push({ range: `'${CP_LEAD_SHEET}'!K${rowIndex}`, values: [[formatted]] });
        }
        if (whatsappNote) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!T${rowIndex}`,
            values: [[(remarks + whatsappNote).trim()]],
          });
        }
        break;

      case "No Connection Yet":
        if (whatsappNote) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!T${rowIndex}`,
            values: [[(remarks + whatsappNote).trim()]],
          });
        }
        break;

      case "Not Qualified":
        if (notQualifiedReason.trim()) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!U${rowIndex}`,
            values: [[notQualifiedReason]],
          });
        }
        break;

      case "Not Interested":
        // No extra fields
        break;

      default:
        return res.status(400).json({
          success: false,
          error: "Invalid status value",
        });
    }

    // Execute batch update
    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: CP_SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates.map((u) => ({
          range: u.range,
          majorDimension: "ROWS",
          values: u.values,
        })),
      },
    });

    res.json({
      success: true,
      message: "Lead updated successfully",
      newFollowUpCount,
    });
  } catch (error) {
    console.error("❌ Update failed:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

module.exports = router;