require("dotenv").config();
console.log("EMAIL:", process.env.GOOGLE_CLIENT_EMAIL);
console.log("KEY EXISTS:", !!process.env.GOOGLE_PRIVATE_KEY);
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

// ============================================
// Use Routes
// ============================================
app.use("/api/leads", nbdinRoutes);

// ============================================
// Health Check Route
// ============================================
app.get("/", (req, res) => {
  res.json({
    message: "🚀 Backend Server is Running!",
    status: "OK",
    timestamp: new Date().toISOString(),
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
});
