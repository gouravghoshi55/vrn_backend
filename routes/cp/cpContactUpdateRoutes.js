// routes/cp/cpContactUpdateRoutes.js
const express = require('express');
const router = express.Router();
const authMiddleware = require('../authRoutes'); // tumhara path jo fix ho gaya

router.post('/update-contact-by-uid', authMiddleware, async (req, res) => {
  try {
    const { uniqueId, newContact } = req.body;

    // Required fields
    if (!uniqueId || newContact === undefined) {
      return res.status(400).json({
        success: false,
        message: 'uniqueId aur newContact zaroori hain',
      });
    }

    // Contact validation: 10 digits, starts with 6-9
    const contactStr = String(newContact).trim();
    if (!/^\d{10}$/.test(contactStr)) {
      return res.status(400).json({
        success: false,
        message: 'Invalid contact number. Must be exactly 10 digits.',
      });
    }

    // Use req.sheets (tumhare middleware se aa raha hai)
    const sheets = req.sheets;

    // Source sheet details
    const spreadsheetId = '17NsMDuq_woISO9CJTBh2e5BaZaKcSXkBEoEF6CNlDd0';
    const sheetName = 'Channel Partner FMS';
    const startRow = 8; // data A8 se shuru

    // Get all data from source sheet (starting from row 8)
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A${startRow}:D`, // A se D tak, row 8 se end tak
    });

    const rows = response.data.values || [];

    // Find row with matching Unique ID (Column B = index 1)
    let targetRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const uidInSheet = row[1]; // Column B (0-based index 1)

      if (uidInSheet && String(uidInSheet).trim().toLowerCase() === String(uniqueId).trim().toLowerCase()) {
        targetRowIndex = startRow + i; // actual sheet row number
        break; // first match only
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({
        success: false,
        message: 'Unique ID not found in source sheet',
      });
    }

    // Update contact number in Column D
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!D${targetRowIndex}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[contactStr]],
      },
    });

    res.status(200).json({
      success: true,
      message: 'Contact number successfully updated in source sheet',
      uniqueId,
      newContact: contactStr,
      updatedRow: targetRowIndex,
    });
  } catch (error) {
    console.error('Contact update by UID error:', error);
    res.status(500).json({
      success: false,
      message: 'Failed to update contact number',
      error: error.message,
    });
  }
});

module.exports = router;