const { sheets } = require("../../config/googleAuth");
const { getCurrentTimestamp } = require("../../utils/dateUtils");
const { uploadPdfToDrive } = require("../../services/googleDriveService");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

// GET — Show rows where W (Planned) NOT NULL & X (Actual) IS NULL
exports.getAgreementData = async (req, res) => {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:AD`,
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, idx) => {
      const planned = row[22]; // W
      const actual  = row[23]; // X

      if (planned && planned.toString().trim() !== "" &&
          (!actual || actual.toString().trim() === "")) {
        filtered.push({
          rowNumber: DATA_START_ROW + idx,
          uniqueId: row[1] || "",
          firmName: row[2] || "",
          contact:  row[3] || "",
          locality: row[4] || "",
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
    const { rowNumber, status, dealsIn, contactInOffice, remark } = req.body;

    if (!rowNumber || !status) {
      return res.status(400).json({ success: false, message: "rowNumber and status required" });
    }

    let pdfLink = "";
    if (req.file) {
      const fileName = `Agreement_${rowNumber}_${Date.now()}.pdf`;
      pdfLink = await uploadPdfToDrive(req.file.buffer, fileName);
    }

    const actualValue = status === "Done" ? getCurrentTimestamp() : "";

    // Update X(Actual), Y(Status), Z(Deals In), AA(Contact), AB(PDF), AC(Remark)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!X${rowNumber}:AC${rowNumber}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          actualValue,
          status,
          dealsIn || "",
          contactInOffice || "",
          pdfLink,
          remark || "",
        ]],
      },
    });

    res.json({ success: true, message: "Agreement updated", pdfLink });
  } catch (err) {
    console.error("Agreement POST error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};