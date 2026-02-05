require("dotenv").config();
const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
const app = express();
const PORT = process.env.PORT || 5000;
// ============================================
// Middleware
// ============================================
app.use(cors());
app.use(express.json());
// ============================================
// Google Sheets Setup (ENV Based)
// ============================================
const auth = new google.auth.JWT({
  email: process.env.GOOGLE_CLIENT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
let sheets;
async function initializeGoogleSheets() {
  try {
    await auth.authorize();
    sheets = google.sheets({ version: "v4", auth });
    console.log("✅ Google Sheets connected successfully!");
  } catch (error) {
    console.error("❌ Google Sheets connection failed:", error);
  }
}
// Initialize on startup
initializeGoogleSheets();
// ============================================
// Middleware - Attach sheets to request
// ============================================
app.use((req, res, next) => {
  if (!sheets) {
    return res.status(503).json({
      success: false,
      error: "Google Sheets not connected. Please try again later.",
    });
  }
  req.sheets = sheets;
  next();
});
// ============================================
// Import Routes
// ============================================
const nbdinRoutes = require("./routes/nbdinRoutes");
const fieldVisitRoutes = require("./routes/fieldVisitRoutes");
// ============================================
// Use Routes
// ============================================
app.use("/api/leads", nbdinRoutes);
app.use("/api/field-visit", fieldVisitRoutes);
// app.use("/nbdin/update", nbdinRoutes);
// app.use("/field-visit/update", fieldVisitRoutes);
// ============================================
// Health Check Route
// ============================================
app.get("/", (req, res) => {
  res.json({
    message: "🚀 Backend Server is Running!",
    status: "OK",
    timestamp: new Date().toISOString(),
    endpoints: {
      nbdin: "/api/leads/nbdin",
      fieldVisit: "/api/field-visit/list",
    },
  });
});
app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    message: "Server is healthy",
    sheetsConnected: !!sheets,
  });
});
// ============================================
// Global Error Handler
// ============================================
app.use((err, req, res, next) => {
  console.error("Server Error:", err);
  res.status(500).json({
    success: false,
    error: "Internal Server Error",
    message: err.message,
  });
});
// ============================================
// Start Server
// ============================================
app.listen(PORT, () => {
  console.log(`🚀 Server is running on port ${PORT}`);
  console.log(`📊 NBDIN API: /api/leads/nbdin`);
  console.log(`🏠 Field Visit API: /api/field-visit/list`);
});