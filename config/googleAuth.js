const { google } = require("googleapis");

// ✅ Singleton with promise caching
let _authClient = null;
let _authPromise = null;  // 🔥 Shared promise for concurrent requests
let _sheetsClient = null;
let _driveClient = null;
let _authorizedAt = 0;
const TOKEN_TTL = 50 * 60 * 1000; // 50 min

function createAuth() {
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").replace(
    /\\n/g,
    "\n"
  );

  if (!process.env.GOOGLE_CLIENT_EMAIL || !privateKey) {
    throw new Error(
      "Missing GOOGLE_CLIENT_EMAIL or GOOGLE_PRIVATE_KEY env vars"
    );
  }

  return new google.auth.JWT({
    email: process.env.GOOGLE_CLIENT_EMAIL,
    key: privateKey,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
  });
}

// 🔥 Get authorized auth client (with promise sharing)
async function getAuth() {
  const now = Date.now();
  const isExpired = now - _authorizedAt > TOKEN_TTL;

  // Return cached authorized client if still valid
  if (_authClient && !isExpired) return _authClient;

  // If auth is currently in progress, wait for that promise
  if (_authPromise) return _authPromise;

  // Start new auth flow
  _authPromise = (async () => {
    const t = Date.now();
    console.log("🔐 [AUTH] Authorizing Google JWT...");
    
    const client = createAuth();
    await client.authorize();  // 🔥 THE MISSING PIECE!
    
    _authClient = client;
    _authorizedAt = Date.now();
    
    // Reset cached service clients (they hold old auth ref)
    _sheetsClient = null;
    _driveClient = null;
    
    console.log(`✅ [AUTH] Authorized in ${Date.now() - t}ms`);
    return client;
  })();

  try {
    return await _authPromise;
  } catch (err) {
    console.error("❌ [AUTH] Failed:", err.message);
    _authClient = null;
    _authorizedAt = 0;
    throw err;
  } finally {
    _authPromise = null;
  }
}

async function getSheets() {
  if (_sheetsClient && Date.now() - _authorizedAt < TOKEN_TTL) {
    return _sheetsClient;
  }
  const auth = await getAuth();
  _sheetsClient = google.sheets({ version: "v4", auth });
  return _sheetsClient;
}

async function getDrive() {
  if (_driveClient && Date.now() - _authorizedAt < TOKEN_TTL) {
    return _driveClient;
  }
  const auth = await getAuth();
  _driveClient = google.drive({ version: "v3", auth });
  return _driveClient;
}

// 🔥 KICK OFF AUTH IMMEDIATELY ON MODULE LOAD
// This runs in parallel with Express boot,
// so by the time first request arrives, auth is ready
getAuth().catch((e) =>
  console.error("⚠️ [AUTH] Boot init failed (will retry on request):", e.message)
);

module.exports = { getAuth, getSheets, getDrive };