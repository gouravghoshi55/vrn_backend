const { drive } = require("../config/googleAuth");
const { Readable } = require("stream");

// ✅ Naya Folder ID update kar diya
const FOLDER_ID = "0ACYEbp3qizyBUk9PVA";

async function uploadPdfToDrive(fileBuffer, fileName) {
  try {
    console.log("📤 Uploading to Drive Folder:", FOLDER_ID);
    console.log("File Name:", fileName);

    const bufferStream = new Readable();
    bufferStream.push(fileBuffer);
    bufferStream.push(null);

    const response = await drive.files.create({
      requestBody: {
        name: fileName,
        parents: [FOLDER_ID],
        mimeType: "application/pdf",
      },
      media: {
        mimeType: "application/pdf",
        body: bufferStream,
      },
      fields: "id, webViewLink",
      supportsAllDrives: true,   // ✅ Shared Drive support ke liye zaroori
    });

    console.log("✅ File uploaded successfully. ID:", response.data.id);

    // Link ko public (anyone with link can view) banane ke liye permission add kar rahe hain
    await drive.permissions.create({
      fileId: response.data.id,
      requestBody: { role: "reader", type: "anyone" },
      supportsAllDrives: true,   // ✅
    });

    console.log("✅ Shareable link generated:", response.data.webViewLink);
    return response.data.webViewLink;

  } catch (err) {
    console.error("❌ Drive upload error:", err.message);
    console.error("Full error details:", err.errors || err);
    throw err;
  }
}

module.exports = { uploadPdfToDrive };