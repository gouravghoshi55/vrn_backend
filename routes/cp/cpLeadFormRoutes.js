const express = require("express");
const router = express.Router();

const CP_SPREADSHEET_ID = process.env.CP_SPREADSHEET_ID;
const CP_LEAD_SHEET = "Channel Partner FMS";
const DATA_START_ROW = 8;

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
  if (!input || !input.includes("T")) return "";
  const [datePart, timePart] = input.split("T");
  const [year, month, day] = datePart.split("-");
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
// FINAL STATUSES - These close the lead
// ============================================
const FINAL_STATUSES = ["Qualified", "Not Interested", "Not Qualified"];

// ============================================
// GET /list - Fetch pending leads
// ============================================

router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching CP Lead Form pending records...");

    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: CP_SPREADSHEET_ID,
      range: `'${CP_LEAD_SHEET}'!A${DATA_START_ROW}:V`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      const uniqueId           = row[1]  || "";
      const customerName       = row[2]  || "";
      const customerContact    = row[3]  || "";
      const interestedIn       = row[4]  || "";
      const leadGenBy          = row[5]  || "";
      const leadGenNumber      = row[6]  || "";
      const leadGenName        = row[7]  || "";
      const leadRemark         = row[8]  || "";
      const followUpCountStr   = row[9]  || "0";
      const planned            = row[10] || "";
      const actual             = row[11] || "";
      const status             = row[12] || "";
      const projectSelection   = row[13] || "";
      const importantNote      = row[14] || "";
      const purpose            = row[15] || "";
      const nextFollowUp       = row[16] || "";
      const sendWhatsapp       = row[17] || "No";
      const plannedSiteVisit   = row[18] || "";
      const remark             = row[19] || "";
      const notQualifiedReason = row[20] || "";
      const canContact         = row[21] || "No";

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
      plannedSiteVisit = "",
      sendWhatsapp = "No",
      whatsappProject = "",
      alternateWhatsapp = "",
      nextFollowUp = "",
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

    // ============================================
    // Check if this is a FINAL status
    // ============================================
    const isFinalStatus = FINAL_STATUSES.includes(status);

    console.log(`📌 Status Type: ${isFinalStatus ? "FINAL (will set Actual)" : "INTERMEDIATE (no Actual)"}`);

    // ============================================
    // Always update these columns
    // ============================================
    
    // J - FollowUp Count (always increment)
    updates.push({
      range: `'${CP_LEAD_SHEET}'!J${rowIndex}`,
      values: [[newFollowUpCount]],
    });

    // M - Status (always update)
    updates.push({
      range: `'${CP_LEAD_SHEET}'!M${rowIndex}`,
      values: [[status]],
    });

    // R - Send Details on WhatsApp (always update)
    updates.push({
      range: `'${CP_LEAD_SHEET}'!R${rowIndex}`,
      values: [[sendWhatsapp]],
    });

    // ============================================
    // L - Actual Date (ONLY for final statuses)
    // ============================================
    if (isFinalStatus) {
      updates.push({
        range: `'${CP_LEAD_SHEET}'!L${rowIndex}`,
        values: [[timestamp]],
      });
      console.log(`✅ Setting Actual date: ${timestamp}`);
    } else {
      console.log(`⏭️ Skipping Actual date (intermediate status)`);
    }

    // ============================================
    // Remarks (T) - always update if provided
    // ============================================
    let finalRemarks = remarks.trim();

    // Add WhatsApp note to remarks if applicable
    if (sendWhatsapp === "Yes" && whatsappProject) {
      const altNum = alternateWhatsapp.trim() || "Same as main number";
      const whatsappNote = `WhatsApp Sent: ${whatsappProject} (Alt: ${altNum})`;
      finalRemarks = finalRemarks ? `${finalRemarks}\n${whatsappNote}` : whatsappNote;
    }

    if (finalRemarks) {
      updates.push({
        range: `'${CP_LEAD_SHEET}'!T${rowIndex}`,
        values: [[finalRemarks]],
      });
    }

    // ============================================
    // Status-specific updates
    // ============================================

    switch (status) {
      case "Qualified":
        // Project Selection (N)
        if (projectSelection) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!N${rowIndex}`,
            values: [[projectSelection]],
          });
        }

        // Important Note (O)
        if (importantNote) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!O${rowIndex}`,
            values: [[importantNote]],
          });
        }

        // Purpose (P)
        if (purpose) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!P${rowIndex}`,
            values: [[purpose]],
          });
        }

        // Planned Site Visit Date (S)
        if (plannedSiteVisit) {
          const formatted = formatDateTimeForSheet(plannedSiteVisit);
          updates.push({
            range: `'${CP_LEAD_SHEET}'!S${rowIndex}`,
            values: [[formatted]],
          });
        }

        // Can Contact (V)
        if (canContact) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!V${rowIndex}`,
            values: [[canContact]],
          });
        }
        break;

      case "Next Followup Required":
        // Next FollowUp Date (Q)
        if (nextFollowUp) {
          const formatted = formatDateTimeForSheet(nextFollowUp);
          updates.push({
            range: `'${CP_LEAD_SHEET}'!Q${rowIndex}`,
            values: [[formatted]],
          });

          // Update Planned (K) to next follow-up date
          // So lead appears again on that date
          updates.push({
            range: `'${CP_LEAD_SHEET}'!K${rowIndex}`,
            values: [[formatted]],
          });
          console.log(`📅 Updated Planned date to: ${formatted}`);
        }
        break;

      case "No Connection Yet":
        // No specific fields, just WhatsApp if sent
        // Lead stays in pending list with same planned date
        break;

      case "Not Qualified":
        // Not Qualified Reason (U)
        if (notQualifiedReason.trim()) {
          updates.push({
            range: `'${CP_LEAD_SHEET}'!U${rowIndex}`,
            values: [[notQualifiedReason]],
          });
        }
        break;

      case "Not Interested":
        // No extra fields needed
        // Actual date is set (handled above)
        break;

      default:
        return res.status(400).json({
          success: false,
          error: "Invalid status value",
        });
    }

    // ============================================
    // Execute batch update
    // ============================================
    console.log(`📤 Executing ${updates.length} cell updates...`);

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

    console.log(`✅ Lead updated successfully!`);

    res.json({
      success: true,
      message: isFinalStatus 
        ? "Lead closed successfully" 
        : "Lead updated, will appear in next followup",
      newFollowUpCount,
      isFinalStatus,
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