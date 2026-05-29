const { sheets } = require("../../config/googleAuth");
const { getCurrentTimestamp } = require("../../utils/dateUtils");
const { uploadPdfToDrive } = require("../../services/googleDriveService");

const SPREADSHEET_ID = "1iGI-DvLlBPj5mmwgOCs926xtaYVgTtoYcD8h2qhhhQc";
const SHEET_NAME = "FMS";
const DATA_START_ROW = 8;

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
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${DATA_START_ROW}:W`,
      valueRenderOption: "FORMATTED_VALUE", // ensure formula results come as text
      dateTimeRenderOption: "FORMATTED_STRING",
    });

    const rows = response.data.values || [];
    const filtered = [];

    rows.forEach((row, idx) => {
      // ✅ Safely read cells (avoid undefined from trailing empty cells)
      const plannedP = (row[15] || "").toString().trim(); // P
      const actual = (row[16] || "").toString().trim(); // Q
      const status = (row[17] || "").toString().trim(); // R
      const nextFollowW = (row[22] || "").toString().trim(); // W

      // const rowNum = DATA_START_ROW + idx;
      // console.log(`Row ${rowNum}:`, {
      //   P: row[15],
      //   Q: row[16],
      //   R: row[17],
      //   W: row[22],
      //   rowLength: row.length,
      // });

      // ✅ Decide effective planned date
      let effectivePlanned = "";
      if (status === "Next Followup Required") {
        // Prefer W; fallback to P if W not yet filled
        effectivePlanned = nextFollowW || plannedP;
      } else {
        effectivePlanned = plannedP;
      }

      // ✅ Show if planned exists AND actual empty
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
    const {
      rowNumber,
      status,
      dealsIn,
      contactInOffice,
      remark,
      nextPlannedDate,
    } = req.body;

    if (!rowNumber || !status) {
      return res
        .status(400)
        .json({ success: false, message: "rowNumber and status required" });
    }

    // ✅ Next Followup Required → write to W (not P)
    if (status === "Next Followup Required") {
      if (!nextPlannedDate) {
        return res.status(400).json({
          success: false,
          message: "nextPlannedDate is required for Next Followup",
        });
      }

      const formattedDate = formatPlannedDateTime(nextPlannedDate);

      await sheets.spreadsheets.values.batchUpdate({
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

    // ✅ Done / other status flow
    let pdfLink = "";
    if (req.file) {
      const fileName = `Agreement_${rowNumber}_${Date.now()}.pdf`;
      pdfLink = await uploadPdfToDrive(req.file.buffer, fileName);
    }

    const actualValue = status === "Done" ? getCurrentTimestamp() : "";

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
