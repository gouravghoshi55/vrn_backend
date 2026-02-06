const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = {
  END_USER: "END USER LEADS FMS",
  CHANNEL_PARTNER: "Channel Partener Lead FMS",
};

// ============================================
// Column Reference (After Field Visit)
// ============================================
// W (22) = Planned
// X (23) = Actual
// Y (24) = Status
// Z (25) = Deal Meeting Date
// AA (26) = Next FollowUp Date
// AB (27) = FollowUp Count
// AC (28) = Remarks

// ============================================
// Helper Functions
// ============================================
async function getFilteredLeads(sheets, sheetName) {
  try {
    // A8:AC tak data fetch
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${sheetName}'!A8:AC`,
    });

    const rows = response.data.values || [];
    const filteredLeads = [];

    rows.forEach((row, index) => {
      // W = index 22 (Planned), X = index 23 (Actual), AB = index 27 (FollowUp Count)
      const plannedDate = row[22] ? row[22].trim() : "";
      const actualDate = row[23] ? row[23].trim() : "";
      const followUpCount = row[27] ? row[27].trim() : "0";

      // Condition: Planned (W) NOT NULL and Actual (X) NULL
      if (plannedDate && !actualDate) {
        filteredLeads.push({
          rowIndex: index + 8,
          sheetName: sheetName,
          uniqueId: row[1] || "",
          customerName: row[2] || "",
          customerContact: row[3] || "",
          interestedIn: row[4] || "",
          projectSelection: row[5] || "",
          leadSource: row[6] || "",
          leadGenNumber: row[7] || "",
          leadGenName: row[8] || "",
          plannedDate: plannedDate,
          followUpCount: parseInt(followUpCount) || 0,
        });
      }
    });

    return filteredLeads;
  } catch (error) {
    console.error(`Error fetching ${sheetName}:`, error.message);
    throw error;
  }
}

function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

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

// ============================================
// Routes
// ============================================
router.get("/list", async (req, res) => {
  try {
    console.log("📊 Fetching After Field Visit data...");
    console.log("   Condition: Planned (W) NOT NULL, Actual (X) NULL");

    const [endUserLeads, channelPartnerLeads] = await Promise.all([
      getFilteredLeads(req.sheets, SHEETS.END_USER),
      getFilteredLeads(req.sheets, SHEETS.CHANNEL_PARTNER),
    ]);

    console.log(`   End User Leads: ${endUserLeads.length}`);
    console.log(`   Channel Partner Leads: ${channelPartnerLeads.length}`);

    let allLeads = [...endUserLeads, ...channelPartnerLeads];

    allLeads.sort((a, b) => {
      const dateA = parseDate(a.plannedDate);
      const dateB = parseDate(b.plannedDate);
      return dateA - dateB;
    });

    console.log(`✅ Total After Field Visit Leads: ${allLeads.length}`);

    res.json({
      success: true,
      data: allLeads,
      total: allLeads.length,
    });
  } catch (error) {
    console.error("❌ Error fetching After Field Visit leads:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to fetch leads",
      message: error.message,
    });
  }
});

router.post("/update", async (req, res) => {
  try {
    const {
      sheetName,
      rowIndex,
      status,
      dealMeetingDate,
      nextFollowUpDate,
      remarks,
      currentFollowUpCount,
    } = req.body;

    console.log("📝 Updating After Field Visit record:", { sheetName, rowIndex, status });

    if (!sheetName || !rowIndex || !status) {
      return res.status(400).json({
        success: false,
        error: "Missing required fields: sheetName, rowIndex, status",
      });
    }

    const timestamp = getCurrentTimestamp();
    const updates = [];

    // Status = "Done" -> Record complete, no more followups
    // Status = "No Conversation" -> FollowUp Count +1, Next FollowUp Date updates Planned
    // Status = "Not Interested" -> Record complete

    if (status === "No Conversation") {
      // FollowUp Count +1
      const newFollowUpCount = (parseInt(currentFollowUpCount) || 0) + 1;

      // W (Planned) = Next FollowUp Date (override)
      if (nextFollowUpDate) {
        updates.push({
          range: `'${sheetName}'!W${rowIndex}`,
          values: [[nextFollowUpDate]],
        });
      }

      // X (Actual) = timestamp
      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[timestamp]],
      });

      // Y (Status) = "No Conversation"
      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [[status]],
      });

      // AA (Next FollowUp Date)
      if (nextFollowUpDate) {
        updates.push({
          range: `'${sheetName}'!AA${rowIndex}`,
          values: [[nextFollowUpDate]],
        });
      }

      // AB (FollowUp Count) = +1
      updates.push({
        range: `'${sheetName}'!AB${rowIndex}`,
        values: [[newFollowUpCount.toString()]],
      });

      // AC (Remarks)
      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[remarks]],
        });
      }

      console.log(`   FollowUp Count: ${newFollowUpCount}`);
    } else if (status === "Done") {
      // X (Actual) = timestamp
      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[timestamp]],
      });

      // Y (Status) = "Done"
      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [[status]],
      });

      // Z (Deal Meeting Date)
      if (dealMeetingDate) {
        updates.push({
          range: `'${sheetName}'!Z${rowIndex}`,
          values: [[dealMeetingDate]],
        });
      }

      // AC (Remarks)
      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[remarks]],
        });
      }
    } else if (status === "Not Interested") {
      // X (Actual) = timestamp
      updates.push({
        range: `'${sheetName}'!X${rowIndex}`,
        values: [[timestamp]],
      });

      // Y (Status) = "Not Interested"
      updates.push({
        range: `'${sheetName}'!Y${rowIndex}`,
        values: [[status]],
      });

      // AC (Remarks)
      if (remarks) {
        updates.push({
          range: `'${sheetName}'!AC${rowIndex}`,
          values: [[remarks]],
        });
      }
    }

    await req.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    console.log(`✅ Updated After Field Visit row ${rowIndex} in ${sheetName}`);
    console.log(`   Status: ${status}`);

    res.json({
      success: true,
      message: "After Field Visit updated successfully",
      data: {
        sheetName,
        rowIndex,
        status,
        dealMeetingDate,
        nextFollowUpDate,
        remarks,
        actualTimestamp: timestamp,
      },
    });
  } catch (error) {
    console.error("❌ Error updating After Field Visit lead:", error.message);
    res.status(500).json({
      success: false,
      error: "Failed to update lead",
      message: error.message,
    });
  }
});

// Debug Route
router.get("/debug", async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'END USER LEADS FMS'!A7:AC15`,
    });

    res.json({
      rawData: response.data.values,
      message: "Row 7 = Headers, Row 8+ = Data",
      columnMapping: {
        "W (22)": "Planned",
        "X (23)": "Actual",
        "Y (24)": "Status",
        "Z (25)": "Deal Meeting Date",
        "AA (26)": "Next FollowUp Date",
        "AB (27)": "FollowUp Count",
        "AC (28)": "Remarks",
      },
    });
  } catch (error) {
    res.json({ error: error.message });
  }
});

module.exports = router;
