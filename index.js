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
// Google Sheets Setup
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
const authRoutes = require("./routes/authRoutes");
const nbdinRoutes = require("./routes/nbdApi/nbdinRoutes");
const fieldVisitRoutes = require("./routes/nbdApi/fieldVisitRoutes");
const afterFieldVisitRoutes = require("./routes/nbdApi/afterFieldVisitRoutes");
const meetingNbdRoutes = require("./routes/nbdApi/meetingNbdRoutes");
const bookingNbdRoutes = require("./routes/nbdApi/bookingNbdRoutes");

// ============================================
// Use Routes
// ============================================
app.use("/api/auth", authRoutes); // ✅ NEW - Auth routes

// Existing routes (will protect these later)
app.use("/api/leads", nbdinRoutes);
app.use("/api/field-visit", fieldVisitRoutes);
app.use("/api/after-field-visit", afterFieldVisitRoutes);
app.use("/api/meeting-nbd", meetingNbdRoutes);
app.use("/api/booking-nbd", bookingNbdRoutes);

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
  console.log(`🔐 Auth API: /api/auth/login`);
  console.log(`📊 NBD API: /api/leads/nbdin`);
  console.log(`📊 Field Visit API: /api/field-visit/list`);
  console.log(`📊 After Field Visit API: /api/after-field-visit/list`);
  console.log(`📊 Meeting NBD API: /api/meeting-nbd/list`);
  console.log(`📊 Booking NBD API: /api/booking-nbd/list`);
});