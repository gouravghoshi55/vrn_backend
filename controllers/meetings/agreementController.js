const { getCurrentTimestamp } = require("../../utils/dateUtils");
const { uploadPdfToDrive } = require("../../services/googleDriveService");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

// ✅ NEW — NI logging in main working sheet
const MAIN_SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

function formatPlannedDateTime(dateStr) {
  if (!dateStr) return "";
  if (dateStr.includes("T")) {
    const [dt, t] = dateStr.split("T");
    const [y, m, d] = dt.split("-");
    const time = t.length === 5 ? `${t}:00` : t;
    return `${d}/${m}/${y} ${time}`;
  }
  return dateStr;
}

// GET — Show pending rows (unchanged)
exports.getAgreementData = async (req, res) => {
  try {
    const response = await req.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:W`,
      valueRenderOption: "FORMATTED_VALUE",
      dateTimeRenderOption: "FORMATTED_STRING",
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, idx) => {
      const plannedP = (row[15] || "").toString().trim();
      const actual = (row[16] || "").toString().trim();
      const status = (row[17] || "").toString().trim();
      const nextFollowW = (row[22] || "").toString().trim();

      let effectivePlanned = "";
      if (status === "Next Followup Required") {
        effectivePlanned = nextFollowW || plannedP;
      } else {
        effectivePlanned = plannedP;
      }

      if (effectivePlanned !== "" && actual === "") {
        filtered.push({
          rowNumber: DATA_START_ROW + idx,
          uniqueId: row[1] || "",
          firmName: row[2] || "",
          contact: row[3] || "",
          locality: row[4] || "",
          plannedDate: effectivePlanned,
          status: status,
        });
      }
    });

    res.json({ success: true, data: filtered });
  } catch (err) {
    console.error("Agreement GET error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};

// POST — Save Agreement action
exports.submitAgreementAction = async (req, res) => {
  try {
    console.log("📄 [Agreement] req.body:", req.body);
    console.log(
      "📄 [Agreement] req.file:",
      req.file
        ? {
            fieldname: req.file.fieldname,
            originalname: req.file.originalname,
            size: req.file.size,
            mimetype: req.file.mimetype,
          }
        : "NO FILE",
    );

    const {
      rowNumber,
      status,
      dealsIn,
      contactInOffice,
      remark,
      nextPlannedDate,
      notInterestedReason, // ✅ NEW
      leadInfo,            // ✅ NEW (may come as string when FormData)
    } = req.body;

    if (!rowNumber || !status) {
      return res
        .status(400)
        .json({ success: false, message: "rowNumber and status required" });
    }

    // ✅ Parse leadInfo if sent as JSON string (FormData scenario)
    let parsedLeadInfo = {};
    try {
      parsedLeadInfo =
        typeof leadInfo === "string" ? JSON.parse(leadInfo) : leadInfo || {};
    } catch {
      parsedLeadInfo = {};
    }

    // ============================================
    // ✅ NEXT FOLLOWUP REQUIRED FLOW
    // ============================================
    if (status === "Next Followup Required") {
      if (!nextPlannedDate) {
        return res.status(400).json({
          success: false,
          message: "nextPlannedDate is required for Next Followup",
        });
      }

      const formattedDate = formatPlannedDateTime(nextPlannedDate);

      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            {
              range: `${SHEET_NAME}!Q${rowNumber}:R${rowNumber}`,
              values: [["", "Next Followup Required"]],
            },
            {
              range: `${SHEET_NAME}!V${rowNumber}`,
              values: [[remark || ""]],
            },
            {
              range: `${SHEET_NAME}!W${rowNumber}`,
              values: [[formattedDate]],
            },
          ],
        },
      });

      return res.json({
        success: true,
        message: "Next Followup scheduled successfully",
      });
    }

    // ============================================
    // ✅ NOT INTERESTED FLOW (NEW)
    // ============================================
    if (status === "Not Interested") {
      if (!notInterestedReason?.trim()) {
        return res.status(400).json({
          success: false,
          message: "Reason required for Not Interested",
        });
      }

      const timestamp = getCurrentTimestamp();

      // Build final remark with reason
      const finalRemark = `${remark || ""}${remark ? " | " : ""}Reason: ${notInterestedReason}`;

      // Update Q (timestamp), R (status), V (remark+reason)
      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            {
              range: `${SHEET_NAME}!Q${rowNumber}:R${rowNumber}`,
              values: [[timestamp, "Not Interested"]],
            },
            {
              range: `${SHEET_NAME}!V${rowNumber}`,
              values: [[finalRemark]],
            },
          ],
        },
      });

      // ✅ Log to Not Interested Reasons sheet
      try {
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
                "Agreement - Not Interested",
                parsedLeadInfo.uniqueId || "",
                parsedLeadInfo.firmName || "",
                parsedLeadInfo.contact || "",
                parsedLeadInfo.locality || "",
                "", // project
                "", // leadSource
                "", // CP name (N/A)
                notInterestedReason,
                req.user?.email || "",
              ],
            ],
          },
        });
        console.log(
          `✅ Agreement NI logged: ${parsedLeadInfo.uniqueId} - ${parsedLeadInfo.firmName}`,
        );
      } catch (logErr) {
        console.warn("⚠️ NI log failed:", logErr.message);
      }

      return res.json({
        success: true,
        message: "Marked Not Interested — logged to NI sheet",
      });
    }

    // ============================================
    // ✅ DONE FLOW (existing — with PDF)
    // ============================================
    let pdfLink = "";
    if (req.file && req.file.buffer) {
      const fileName = `Agreement_${rowNumber}_${Date.now()}.pdf`;
      try {
        pdfLink = await uploadPdfToDrive(req.file.buffer, fileName);
        console.log("✅ PDF uploaded successfully:", pdfLink);
      } catch (uploadErr) {
        console.error("❌ PDF upload failed:", uploadErr.message);
      }
    }

    const actualValue = status === "Done" ? getCurrentTimestamp() : "";

    await req.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!Q${rowNumber}:V${rowNumber}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [
          [
            actualValue,
            status,
            dealsIn || "",
            contactInOffice || "",
            pdfLink,
            remark || "",
          ],
        ],
      },
    });

    res.json({ success: true, message: "Agreement updated", pdfLink });
  } catch (err) {
    console.error("Agreement POST error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};