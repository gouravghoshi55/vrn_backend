const { sheets } = require("../../config/googleAuth");
const { getCurrentTimestamp } = require("../../utils/dateUtils");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

exports.getFullKittingData = async (req, res) => {
  try {
    console.log("\n========== FULL KITTING API HIT ==========");
    console.log("Spreadsheet ID:", SPREADSHEET_ID);
    console.log("Sheet Name:", SHEET_NAME);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:N`,
    });

    const rows = response.data.values || [];
    console.log("Total rows fetched:", rows.length);

    if (rows.length > 0) {
      console.log("First row data:", JSON.stringify(rows[0]));
      console.log("Row[0] length:", rows[0].length);
      console.log("Column H (index 7) =", rows[0][7]);
      console.log("Column I (index 8) =", rows[0][8]);
    }

    const filtered = [];

    rows.forEach((row, idx) => {
      const planned = row[7];
      const actual  = row[8];

      console.log(`Row ${DATA_START_ROW + idx}: H="${planned}" | I="${actual}"`);

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

    console.log("✅ Filtered count:", filtered.length);
    console.log("==========================================\n");

    res.json({ success: true, data: filtered });
  } catch (err) {
    console.error("❌ FullKitting GET error:", err.message);
    console.error("Full error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};

exports.submitFullKittingAction = async (req, res) => {
  try {
    const { rowNumber, status, pptPrint, slabStructure, remark } = req.body;
    if (!rowNumber || !status) {
      return res.status(400).json({ success: false, message: "rowNumber and status required" });
    }
    const actualValue = status === "Done" ? getCurrentTimestamp() : "";
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!I${rowNumber}:M${rowNumber}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[actualValue, status, pptPrint || "", slabStructure || "", remark || ""]],
      },
    });
    res.json({ success: true, message: "Full Kitting updated" });
  } catch (err) {
    console.error("FullKitting POST error:", err);
    res.status(500).json({ success: false, message: err.message });
  }
};