// Returns timestamp in DD/MM/YYYY HH:mm:ss format (matching your sheet)
function getCurrentTimestamp() {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yyyy = now.getFullYear();
  const hh = String(now.getHours()).padStart(2, "0");
  const min = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return `${dd}/${mm}/${yyyy} ${hh}:${min}:${ss}`;
}

// Check if date is today
function isToday(dateValue) {
  if (!dateValue) return false;
  const d = parseSheetDate(dateValue);
  if (!d) return false;
  const today = new Date();
  return (
    d.getDate() === today.getDate() &&
    d.getMonth() === today.getMonth() &&
    d.getFullYear() === today.getFullYear()
  );
}

// Parse DD/MM/YYYY or DD/MM/YYYY HH:mm:ss
function parseSheetDate(value) {
  if (!value) return null;
  const str = value.toString().trim();
  const parts = str.split(" ")[0].split("/");
  if (parts.length !== 3) return new Date(str);
  const [dd, mm, yyyy] = parts;
  return new Date(`${yyyy}-${mm}-${dd}`);
}

module.exports = { getCurrentTimestamp, isToday, parseSheetDate };