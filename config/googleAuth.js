const { google } = require("googleapis");

// ✅ Singleton — persists within warm container
let _authClient = null;
let _sheetsClient = null;
let _driveClient = null;

function getAuth() {
  if (_authClient) return _authClient;

  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").replace(
    /\\n/g,
    "\n"
  );

  if (!process.env.GOOGLE_CLIENT_EMAIL || !privateKey) {
    throw new Error(
      "Missing GOOGLE_CLIENT_EMAIL or GOOGLE_PRIVATE_KEY env vars"
    );
  }

  _authClient = new google.auth.JWT({
    email: process.env.GOOGLE_CLIENT_EMAIL,
    key: privateKey,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
  });

  return _authClient;
}

function getSheets() {
  if (_sheetsClient) return _sheetsClient;
  _sheetsClient = google.sheets({ version: "v4", auth: getAuth() });
  return _sheetsClient;
}

function getDrive() {
  if (_driveClient) return _driveClient;
  _driveClient = google.drive({ version: "v3", auth: getAuth() });
  return _driveClient;
}

// ✅ Backward compatibility — Proxy makes `sheets.spreadsheets.values.get()` work
// without changing any existing route code
const sheets = new Proxy(
  {},
  {
    get(_, prop) {
      return getSheets()[prop];
    },
  }
);

const drive = new Proxy(
  {},
  {
    get(_, prop) {
      return getDrive()[prop];
    },
  }
);

const auth = new Proxy(
  {},
  {
    get(_, prop) {
      return getAuth()[prop];
    },
  }
);

module.exports = { auth, sheets, drive, getAuth, getSheets, getDrive };