const jwt = require("jsonwebtoken");

// ============================================
// Verify JWT Token
// ============================================
exports.protect = (req, res, next) => {
  try {
    const token = req.headers.authorization?.replace("Bearer ", "");

    if (!token) {
      return res.status(401).json({
        success: false,
        message: "Not authorized, no token",
      });
    }

    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    req.user = decoded; // Attach user info to request
    next();
  } catch (error) {
    res.status(401).json({
      success: false,
      message: "Not authorized, token failed",
    });
  }
};

// ============================================
// Check Module Access
// ============================================
exports.checkModuleAccess = (requiredModule) => {
  return (req, res, next) => {
    const { assignedModule, role } = req.user;

    // Admin has access to all modules
    if (role === "admin" || assignedModule === "all") {
      return next();
    }

    // Check if user has access to required module
    if (assignedModule !== requiredModule) {
      return res.status(403).json({
        success: false,
        message: `Access denied. You don't have permission to access ${requiredModule} module.`,
        allowedModule: assignedModule,
      });
    }

    next();
  };
};