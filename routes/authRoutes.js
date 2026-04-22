const express = require("express");
const router = express.Router();
const jwt = require("jsonwebtoken");

const USER_SPREADSHEET_ID = process.env.USER_SPREADSHEET_ID; // ✅ Changed
const USERS_SHEET = "USERS";

// ============================================
// HELPER - Get Users from Sheet
// ============================================
async function getUsers(sheets) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: USER_SPREADSHEET_ID, // ✅ Changed
    range: `'${USERS_SHEET}'!A2:G`,
  });

  const rows = response.data.values || [];
  return rows.map((row) => ({
    email: row[0]?.trim().toLowerCase() || "",
    password: row[1]?.trim() || "",
    role: row[2]?.trim() || "",
    name: row[3]?.trim() || "",
    assignedModule: row[4]?.trim() || "",
    status: row[5]?.trim() || "active",
    createdAt: row[6] || "",
  }));
}

// ============================================
// LOGIN - POST /api/auth/login
// ============================================
router.post("/login", async (req, res) => {
  try {
    const { email, password } = req.body;

    console.log("Login attempt:", { email });

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        message: "Email and password are required",
      });
    }

    // Get users from sheet
    const users = await getUsers(req.sheets);
    console.log("Total users found:", users.length);
    console.log(
      "All emails:",
      users.map((u) => u.email),
    );

    // Find user
    const user = users.find((u) => u.email === email.toLowerCase().trim());

    if (!user) {
      console.log("User not found:", email);
      return res.status(401).json({
        success: false,
        message: "Invalid email or password",
      });
    }

    // Check if user is active
    if (user.status !== "active") {
      return res.status(403).json({
        success: false,
        message: "Account is inactive. Contact admin.",
      });
    }

    // Check password (plain text comparison)
    if (user.password !== password) {
      console.log("Password mismatch for:", email);
      return res.status(401).json({
        success: false,
        message: "Invalid email or password",
      });
    }

    // Generate JWT Token
    const token = jwt.sign(
      {
        email: user.email,
        role: user.role,
        assignedModule: user.assignedModule,
        name: user.name,
      },
      process.env.JWT_SECRET,
      { expiresIn: process.env.JWT_EXPIRE || "7d" },
    );

    console.log("✅ Login successful:", user.email);

    res.json({
      success: true,
      message: "Login successful",
      token,
      user: {
        email: user.email,
        name: user.name,
        role: user.role,
        assignedModule: user.assignedModule,
      },
    });
  } catch (error) {
    console.error("❌ Login error:", error);
    res.status(500).json({
      success: false,
      message: "Server error",
      error: error.message,
    });
  }
});

// ============================================
// VERIFY TOKEN - GET /api/auth/verify
// ============================================
router.get("/verify", async (req, res) => {
  try {
    const token = req.headers.authorization?.replace("Bearer ", "");

    if (!token) {
      return res.status(401).json({
        success: false,
        message: "No token provided",
      });
    }

    const decoded = jwt.verify(token, process.env.JWT_SECRET);

    res.json({
      success: true,
      user: {
        email: decoded.email,
        name: decoded.name,
        role: decoded.role,
        assignedModule: decoded.assignedModule,
      },
    });
  } catch (error) {
    res.status(401).json({
      success: false,
      message: "Invalid or expired token",
    });
  }
});

// ============================================
// LOGOUT - POST /api/auth/logout
// ============================================
router.post("/logout", (req, res) => {
  res.json({
    success: true,
    message: "Logged out successfully",
  });
});

module.exports = router;
