const { sheets } = require("../../config/googleAuth");
const { getCurrentTimestamp } = require("../../utils/dateUtils");
const { uploadPdfToDrive } = require("../../services/googleDriveService");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

// ✅ Helper — convert "YYYY-MM-DDTHH:mm" to "DD/MM/YYYY HH:mm:ss"
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

// GET — Show rows where P (Planned) NOT NULL & Q (Actual) IS NULL
exports.getAgreementData = async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:V`,
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, idx) => {
      const planned = row[15]; // P
      const actual = row[16];  // Q

      if (
        planned &&
        planned.toString().trim() !== "" &&
        (!actual || actual.toString().trim() === "")
      ) {
        filtered.push({
          rowNumber: DATA_START_ROW + idx,
          uniqueId: row[1] || "",
          firmName: row[2] || "",
          contact: row[3] || "",
          locality: row[4] || "",
          plannedDate: planned,
        });
      }
    });

    res.json({ success: true, data: filtered });
  } catch (err) {
    console.error("Agreement GET error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};

// POST — Save Agreement action (with optional PDF upload)
exports.submitAgreementAction = async (req, res) => {
  try {
    const {
      rowNumber,
      status,
      dealsIn,
      contactInOffice,
      remark,
      nextPlannedDate, // ✅ NEW
    } = req.body;

    if (!rowNumber || !status) {
      return res
        .status(400)
        .json({ success: false, message: "rowNumber and status required" });
    }

    // ✅ Next Followup Required — override Planned (P), no actual, no PDF needed
    if (status === "Next Followup Required") {
      if (!nextPlannedDate) {
        return res.status(400).json({
          success: false,
          message: "nextPlannedDate is required for Next Followup",
        });
      }

      const formattedDate = formatPlannedDateTime(nextPlannedDate);

      // Update P (Planned), Q (Actual=blank), R (Status), V (Remark)
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: [
            {
              range: `${SHEET_NAME}!P${rowNumber}`,
              values: [[formattedDate]],
            },
            {
              range: `${SHEET_NAME}!Q${rowNumber}:R${rowNumber}`,
              values: [["", "Next Followup Required"]],
            },
            {
              range: `${SHEET_NAME}!V${rowNumber}`,
              values: [[remark || ""]],
            },
          ],
        },
      });

      return res.json({
        success: true,
        message: "Next Followup scheduled successfully",
      });
    }

    // ✅ Existing Done flow
    let pdfLink = "";
    if (req.file) {
      const fileName = `Agreement_${rowNumber}_${Date.now()}.pdf`;
      pdfLink = await uploadPdfToDrive(req.file.buffer, fileName);
    }

    const actualValue = status === "Done" ? getCurrentTimestamp() : "";

    // Update Q(Actual), R(Status), S(Deals In), T(Contact), U(PDF), V(Remark)
    await sheets.spreadsheets.values.update({
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