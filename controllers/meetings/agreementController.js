const { getCurrentTimestamp } = require("../../utils/dateUtils");
const { uploadPdfToDrive } = require("../../services/googleDriveService");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

const MAIN_SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const NOT_INTERESTED_SHEET = "Not Interested Reasons";

// ✅ NEW — Short timestamp for remarks
function getShortTimestamp() {
  const now = new Date();
  const d  = String(now.getDate()).padStart(2, "0");
  const mo = String(now.getMonth() + 1).padStart(2, "0");
  const y  = now.getFullYear();
  const h  = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  return `[${d}/${mo}/${y} ${h}:${mi}]`;
}

// ✅ NEW — Append remark with timestamp
async function buildAppendedRemarks(sheets, cellRange, newRemark) {
  if (!newRemark || !String(newRemark).trim()) return null;
  let existing = "";
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!${cellRange}`,
    });
    existing = r.data.values?.[0]?.[0] || "";
  } catch (e) {}
  const timestamped = `${getShortTimestamp()} ${String(newRemark).trim()}`;
  return existing.trim() ? `${existing.trim()}\n${timestamped}` : timestamped;
}

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

// GET — Show pending rows
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
      const plannedP    = (row[15] || "").toString().trim();
      const actual      = (row[16] || "").toString().trim();
      const status      = (row[17] || "").toString().trim();
      const nextFollowW = (row[22] || "").toString().trim();
      const remark      = (row[21] || "").toString(); // ✅ NEW — V column (index 21)

      let effectivePlanned = "";
      if (status === "Next Followup Required") {
        effectivePlanned = nextFollowW || plannedP;
      } else {
        effectivePlanned = plannedP;
      }

      if (effectivePlanned !== "" && actual === "") {
        filtered.push({
          rowNumber:   DATA_START_ROW + idx,
          uniqueId:    row[1] || "",
          firmName:    row[2] || "",
          contact:     row[3] || "",
          locality:    row[4] || "",
          plannedDate: effectivePlanned,
          status:      status,
          remark,  // ✅ NEW
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
    console.log("📄 [Agreement] req.file:", req.file ? {
      fieldname:    req.file.fieldname,
      originalname: req.file.originalname,
      size:         req.file.size,
      mimetype:     req.file.mimetype,
    } : "NO FILE");

    const {
      rowNumber, status, dealsIn, contactInOffice, remark,
      nextPlannedDate, notInterestedReason, leadInfo,
    } = req.body;

    if (!rowNumber || !status) {
      return res
        .status(400)
        .json({ success: false, message: "rowNumber and status required" });
    }

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

      // ✅ Append remark with timestamp
      const appendedRemark = remark && String(remark).trim()
        ? await buildAppendedRemarks(req.sheets, `V${rowNumber}`, remark)
        : null;

      const updateData = [
        {
          range: `${SHEET_NAME}!Q${rowNumber}:R${rowNumber}`,
          values: [["", "Next Followup Required"]],
        },
        {
          range: `${SHEET_NAME}!W${rowNumber}`,
          values: [[formattedDate]],
        },
      ];

      if (appendedRemark !== null) {
        updateData.push({
          range: `${SHEET_NAME}!V${rowNumber}`,
          values: [[appendedRemark]],
        });
      }

      await req.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: "USER_ENTERED", data: updateData },
      });

      return res.json({
        success: true,
        message: "Next Followup scheduled successfully",
      });
    }

    // ============================================
    // ✅ NOT INTERESTED FLOW
    // ============================================
    if (status === "Not Interested") {
      if (!notInterestedReason?.trim()) {
        return res.status(400).json({
          success: false,
          message: "Reason required for Not Interested",
        });
      }

      const timestamp = getCurrentTimestamp();

      // Build remark with reason + append with timestamp
      const remarkWithReason = remark && String(remark).trim()
        ? `${remark.trim()} | Reason: ${notInterestedReason}`
        : `Reason: ${notInterestedReason}`;

      const appendedRemark = await buildAppendedRemarks(
        req.sheets, `V${rowNumber}`, remarkWithReason,
      );

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
              values: [[appendedRemark]],
            },
          ],
        },
      });

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
                "", "", "",
                notInterestedReason,
                req.user?.email || "",
              ],
            ],
          },
        });
        console.log(`✅ Agreement NI logged: ${parsedLeadInfo.uniqueId} - ${parsedLeadInfo.firmName}`);
      } catch (logErr) {
        console.warn("⚠️ NI log failed:", logErr.message);
      }

      return res.json({
        success: true,
        message: "Marked Not Interested — logged to NI sheet",
      });
    }

    // ============================================
    // ✅ DONE FLOW (with PDF)
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

    // ✅ Append remark with timestamp
    let finalRemark = "";
    if (remark && String(remark).trim()) {
      finalRemark = await buildAppendedRemarks(req.sheets, `V${rowNumber}`, remark);
    } else {
      // Keep existing
      try {
        const r = await req.sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: `${SHEET_NAME}!V${rowNumber}`,
        });
        finalRemark = r.data.values?.[0]?.[0] || "";
      } catch (e) {
        finalRemark = "";
      }
    }

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
            finalRemark,
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