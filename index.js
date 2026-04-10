require("dotenv").config();
const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
const app = express();
const PORT = process.env.PORT || 5000;

// ============================================
// Middleware
// ============================================

// 1. CORS
app.use(cors({
  origin: "https://vrn-sales.vercel.app",
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
}));

// 2. Handle OPTIONS Preflight
app.use((req, res, next) => {
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', 'https://vrn-sales.vercel.app');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    return res.status(200).end();
  }
  next();
});

// 3. Body Parsing
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
const { protect } = require("./middleware/authMiddleware");

const authRoutes = require("./routes/authRoutes");
const nbdinRoutes = require("./routes/nbdApi/nbdinRoutes");
const nbdFieldVisitRoutes = require("./routes/nbdApi/fieldVisitRoutes");
const nbdAfterFieldVisitRoutes = require("./routes/nbdApi/afterFieldVisitRoutes");
const nbdMeetingRoutes = require("./routes/nbdApi/meetingNbdRoutes");
const cpFollowupRoutes = require("./routes/cp/cpFollowupRoutes");
const cpFieldVisitRoutes = require("./routes/cp/cpFieldVisitRoutes");
const cpAfterFieldVisitRoutes = require("./routes/cp/cpAfterFieldVisitRoutes");
const cpMeetingRoutes = require("./routes/cp/cpMeetingRoutes");
const cpBookingRoutes = require("./routes/cp/cpBookingRoutes");
const cpLeadFormRoutes = require("./routes/cp/cpLeadFormRoutes");
const cpContactUpdateRoutes = require("./routes/cp/cpContactUpdateRoutes");
const leadSearchRoutes = require("./routes/leadSearch");

// ============================================
// Use Routes
// ============================================

app.use("/api/auth", authRoutes);
app.use("/api/leads", protect, nbdinRoutes);
app.use("/api/field-visit", protect, nbdFieldVisitRoutes);
app.use("/api/after-field-visit", protect, nbdAfterFieldVisitRoutes);
app.use("/api/meeting-nbd", protect, nbdMeetingRoutes);
app.use("/api/cp/followup", cpFollowupRoutes);
app.use("/api/cp/field-visit", cpFieldVisitRoutes);
app.use("/api/cp/after-field-visit", cpAfterFieldVisitRoutes);
app.use("/api/cp/meeting", cpMeetingRoutes);
app.use("/api/cp/booking", cpBookingRoutes);
app.use("/api/cp/lead-form", cpLeadFormRoutes);
app.use("/cp", cpContactUpdateRoutes);
app.use("/api/leads", protect, leadSearchRoutes);

// ============================================
// Health Check
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
app.listen(PORT, "0.0.0.0", () => {
  console.log(`🚀 Server is running on port ${PORT}`);
});

module.exports = app;