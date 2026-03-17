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

// ✅ Import protect middleware
// If your file path is different, update accordingly
const { protect } = require("./middleware/authMiddleware");

// Auth Routes
const authRoutes = require("./routes/authRoutes");

// NBD Routes (END USER LEADS FMS only)
const nbdinRoutes = require("./routes/nbdApi/nbdinRoutes");
const nbdFieldVisitRoutes = require("./routes/nbdApi/fieldVisitRoutes");
const nbdAfterFieldVisitRoutes = require("./routes/nbdApi/afterFieldVisitRoutes");
const nbdMeetingRoutes = require("./routes/nbdApi/meetingNbdRoutes");
const nbdBookingRoutes = require("./routes/nbdApi/bookingNbdRoutes");

// CP Routes (Channel Partener Lead FMS only)
const cpFollowupRoutes = require("./routes/cp/cpFollowupRoutes");
const cpFieldVisitRoutes = require("./routes/cp/cpFieldVisitRoutes");
const cpAfterFieldVisitRoutes = require("./routes/cp/cpAfterFieldVisitRoutes");
const cpMeetingRoutes = require("./routes/cp/cpMeetingRoutes");
const cpBookingRoutes = require("./routes/cp/cpBookingRoutes");
const cpLeadFormRoutes = require("./routes/cp/cpLeadFormRoutes");
const cpContactUpdateRoutes = require('./routes/cp/cpContactUpdateRoutes');

// Add this with other route imports
const leadSearchRoutes = require("./routes/leadSearch");

// ============================================
// Use Routes
// ============================================

// Auth
app.use("/api/auth", authRoutes);

// ✅ NBD APIs — NOW WITH protect MIDDLEWARE for doer-based filtering
app.use("/api/leads", protect, nbdinRoutes);
app.use("/api/field-visit", protect, nbdFieldVisitRoutes);
app.use("/api/after-field-visit", protect, nbdAfterFieldVisitRoutes);
app.use("/api/meeting-nbd", protect, nbdMeetingRoutes);
app.use("/api/booking-nbd", protect, nbdBookingRoutes);

// CP APIs (BDM2 + Admin)
app.use("/api/cp/followup", cpFollowupRoutes);
app.use("/api/cp/field-visit", cpFieldVisitRoutes);
app.use("/api/cp/after-field-visit", cpAfterFieldVisitRoutes);
app.use("/api/cp/meeting", cpMeetingRoutes);
app.use("/api/cp/booking", cpBookingRoutes);
app.use("/api/cp/lead-form", cpLeadFormRoutes);
app.use('/cp', cpContactUpdateRoutes);

// Add this with other route registrations
app.use("/api/leads", protect, leadSearchRoutes);


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
  console.log(`\n📊 ===== NBD APIs (END USER) =====`);
  console.log(`   Follow-up:         /api/leads/nbdin`);
  console.log(`   Field Visit:       /api/field-visit/list`);
  console.log(`   After Field Visit: /api/after-field-visit/list`);
  console.log(`   Meeting:           /api/meeting-nbd/list`);
  console.log(`   Booking:           /api/booking-nbd/list`);
  console.log(`\n📊 ===== CP APIs (Channel Partner) =====`);
  console.log(`   Follow-up:         /api/cp/followup/can-contact/list`);
  console.log(`   Follow-up:         /api/cp/followup/cannot-contact/list`);
  console.log(`   Field Visit:       /api/cp/field-visit/can-contact/list`);
  console.log(`   Field Visit:       /api/cp/field-visit/cannot-contact/list`);
  console.log(`   After Field Visit: /api/cp/after-field-visit/can-contact/list`);
  console.log(`   After Field Visit: /api/cp/after-field-visit/cannot-contact/list`);
  console.log(`   Meeting:           /api/cp/meeting/can-contact/list`);
  console.log(`   Meeting:           /api/cp/meeting/cannot-contact/list`);
  console.log(`   Booking:           /api/cp/booking/can-contact/list`);
  console.log(`   Booking:           /api/cp/booking/cannot-contact/list`);
  console.log(`   Lead Form:         /api/cp/lead-form/list`);
});