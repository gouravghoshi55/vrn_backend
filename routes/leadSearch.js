const express = require("express");
const router = express.Router();

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = "END USER LEADS FMS";
const CP_SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CP_LEAD_SHEET = "Channel Partner Lead FMS";

// ============================================
// Step Detection Logic
// ============================================

function detectLeadStep(row) {
  // Helper to get column value safely
  const getCol = (index) => (row[index] ? row[index].trim() : "");

  // Basic Lead Info
  const uniqueId = getCol(1);       // B
  const customerName = getCol(2);   // C
  const customerContact = getCol(3); // D
  const projectSelection = getCol(5); // F
  const leadSource = getCol(6);     // G

  // Step 1: First Follow Up - Column O (index 14)
  const step1Status = getCol(14);
  const isInStep1 = step1Status === "" || 
                    step1Status.toLowerCase() === "no conversation" || 
                    step1Status.toLowerCase() === "next follow up";

  // Step 2: Site Visit - Planned on Column U (index 20), Status on Column W (index 22)
  const step2PlannedDate = getCol(20);
  const step2Status = getCol(22);
  const isInStep2 = step2PlannedDate && 
                    (step2Status === "" || 
                     step2Status.toLowerCase() === "pending" ||
                     step2Status.toLowerCase() === "rescheduled");

  // Step 3: After Site Visit Follow Up
  // Planned: AA (26), Actual: AB (27), Status: AC (28)
  const step3PlannedDate = getCol(26);
  const step3ActualDate = getCol(27);
  const step3Status = getCol(28);
  const step3StatusLower = step3Status.toLowerCase();
  const isInStep3 = (step3PlannedDate && !step3ActualDate) || 
                    step3StatusLower === "no conversation" || 
                    step3StatusLower === "next follow up" ||
                    step3StatusLower === "next field visit required";

  // Step 4: Meeting
  // Planned: AH (33), Actual: AI (34), Status: AJ (35)
  const step4PlannedDate = getCol(33);
  const step4ActualDate = getCol(34);
  const step4Status = getCol(35);
  const step4StatusLower = step4Status.toLowerCase();
  const isInStep4 = step4PlannedDate && 
                    (!step4Status || 
                     step4StatusLower === "rescheduled" || 
                     step4StatusLower === "next field visit required");

  // Step 5: Booking
  // Planned: AM (38)
  const step5PlannedDate = getCol(38);
  const isInStep5 = step5PlannedDate !== "";

  // Determine current step (check from highest to lowest)
  let currentStep = null;
  let stepName = "";
  let stepStatus = "";
  let stepPlannedDate = "";
  let stepColor = "";
  let stepIcon = "";

if (isInStep1) {
  currentStep = 1;
  stepName = "First Follow-Up";
  stepStatus = step1Status || "Pending";
  stepPlannedDate = "";
  stepColor = "#3b82f6"; // Blue
  stepIcon = "bi-telephone-fill";
} else if (isInStep2) {
  currentStep = 2;
  stepName = "Site Visit";
  stepStatus = step2Status || "Pending";
  stepPlannedDate = step2PlannedDate;
  stepColor = "#ec4899"; // Pink
  stepIcon = "bi-geo-alt-fill";
} else if (isInStep3) {
  currentStep = 3;
  stepName = "After Site Visit Follow-Up";
  stepStatus = step3Status || "Pending";
  stepPlannedDate = step3PlannedDate;
  stepColor = "#f59e0b"; // Orange
  stepIcon = "bi-chat-dots-fill";
} else if (isInStep4) {
  currentStep = 4;
  stepName = "Meeting";
  stepStatus = step4Status || "Pending";
  stepPlannedDate = step4PlannedDate;
  stepColor = "#8b5cf6"; // Purple
  stepIcon = "bi-calendar-check-fill";
} else if (isInStep5) {
  currentStep = 5;
  stepName = "Booking";
  stepStatus = "Booking Pending";
  stepPlannedDate = step5PlannedDate;
  stepColor = "#10b981"; // Green
  stepIcon = "bi-check-circle-fill";
}

  // Check if lead is completed/closed
  const finalStatuses = ["done", "not interested", "negotiation failed", "deal not done", "booked"];
  
  // Check Step 4 status for completion
  if (step4StatusLower && finalStatuses.includes(step4StatusLower)) {
    currentStep = 0;
    stepName = "Completed/Closed";
    stepStatus = step4Status;
    stepColor = "#6b7280"; // Gray
    stepIcon = "bi-check-all";
  }

  return {
    uniqueId,
    customerName,
    customerContact,
    projectSelection,
    leadSource,
    currentStep,
    stepName,
    stepStatus,
    stepPlannedDate,
    stepColor,
    stepIcon,
  };
}

// ============================================
// Search in Sheet
// ============================================

async function searchInSheet(sheets, spreadsheetId, sheetName, searchTerm, sheetType) {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: `'${sheetName}'!A8:AO`,
    });

    const rows = response.data.values || [];
    const results = [];
    const searchLower = searchTerm.toLowerCase().trim();

    rows.forEach((row, index) => {
      const customerName = row[2] ? row[2].trim().toLowerCase() : "";
      const customerContact = row[3] ? row[3].trim() : "";
      const uniqueId = row[1] ? row[1].trim().toLowerCase() : "";
      const leadGenName = row[8] ? row[8].trim().toLowerCase() : "";

      // Search by name, contact, or unique ID
      const isMatch = 
        customerName.includes(searchLower) ||
        customerContact.includes(searchLower) ||
        uniqueId.includes(searchLower) ||
        leadGenName.includes(searchLower);

      if (isMatch) {
        const leadInfo = detectLeadStep(row);
        
        if (leadInfo.currentStep !== null) {
          results.push({
            ...leadInfo,
            rowIndex: index + 8,
            sheetName: sheetName,
            sheetType: sheetType, // "Direct" or "Channel Partner"
            spreadsheetId: spreadsheetId,
          });
        }
      }
    });

    return results;
  } catch (error) {
    console.error(`Error searching in ${sheetName}:`, error.message);
    return [];
  }
}

// ============================================
// ROUTES
// ============================================

router.get("/search", async (req, res) => {
  try {
    const { q } = req.query;

    if (!q || q.trim().length < 2) {
      return res.status(400).json({
        success: false,
        error: "Search term must be at least 2 characters",
      });
    }

    const searchTerm = q.trim();

    // Search in both sheets
    const [directResults, cpResults] = await Promise.all([
      searchInSheet(req.sheets, SPREADSHEET_ID, SHEET_NAME, searchTerm, "Direct"),
      searchInSheet(req.sheets, CP_SPREADSHEET_ID, CP_LEAD_SHEET, searchTerm, "Channel Partner"),
    ]);

    // Combine results
    const allResults = [...directResults, ...cpResults];

    // Sort by step (highest first) and then by name
    allResults.sort((a, b) => {
      if (b.currentStep !== a.currentStep) {
        return b.currentStep - a.currentStep;
      }
      return a.customerName.localeCompare(b.customerName);
    });

    res.json({
      success: true,
      data: allResults,
      total: allResults.length,
      searchTerm: searchTerm,
    });
  } catch (err) {
    console.error("Search error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Get lead details by unique ID
router.get("/lead/:uniqueId", async (req, res) => {
  try {
    const { uniqueId } = req.params;

    if (!uniqueId) {
      return res.status(400).json({
        success: false,
        error: "Unique ID is required",
      });
    }

    // Search in both sheets
    const [directResults, cpResults] = await Promise.all([
      searchInSheet(req.sheets, SPREADSHEET_ID, SHEET_NAME, uniqueId, "Direct"),
      searchInSheet(req.sheets, CP_SPREADSHEET_ID, CP_LEAD_SHEET, uniqueId, "Channel Partner"),
    ]);

    const allResults = [...directResults, ...cpResults];
    
    // Find exact match
    const exactMatch = allResults.find(
      (r) => r.uniqueId.toLowerCase() === uniqueId.toLowerCase()
    );

    if (exactMatch) {
      res.json({
        success: true,
        data: exactMatch,
      });
    } else {
      res.status(404).json({
        success: false,
        error: "Lead not found",
      });
    }
  } catch (err) {
    console.error("Lead fetch error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;