const { getSheets, getDrive } = require("../config/googleAuth");

const RETRYABLE_CODES = ["ECONNRESET", "ETIMEDOUT", "ENOTFOUND", "EAI_AGAIN"];
const RETRYABLE_MESSAGES = [
  "socket hang up",
  "Client network socket disconnected",
  "ECONNRESET",
];

function isRetryable(err) {
  if (!err) return false;
  if (RETRYABLE_CODES.includes(err.code)) return true;
  if (err.cause && RETRYABLE_CODES.includes(err.cause.code)) return true;
  const msg = err.message || "";
  return RETRYABLE_MESSAGES.some((e) => msg.includes(e));
}

async function withRetry(fn, maxAttempts = 4, baseDelay = 400) {
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;

      if (!isRetryable(err) || attempt === maxAttempts) {
        throw err;
      }

      const delay = baseDelay * Math.pow(2, attempt - 1) + Math.random() * 200;
      console.warn(
        `⚠️ Sheets API retry ${attempt}/${maxAttempts} after ${Math.round(delay)}ms — ${err.code || err.message?.substring(0, 80)}`
      );
      await new Promise((r) => setTimeout(r, delay));
    }
  }
  throw lastErr;
}

function getRetryableSheets() {
  const base = getSheets();
  return {
    spreadsheets: {
      values: {
        get: (p) => withRetry(() => base.spreadsheets.values.get(p)),
        update: (p) => withRetry(() => base.spreadsheets.values.update(p)),
        append: (p) => withRetry(() => base.spreadsheets.values.append(p)),
        batchUpdate: (p) =>
          withRetry(() => base.spreadsheets.values.batchUpdate(p)),
        clear: (p) => withRetry(() => base.spreadsheets.values.clear(p)),
        batchGet: (p) => withRetry(() => base.spreadsheets.values.batchGet(p)),
      },
      get: (p) => withRetry(() => base.spreadsheets.get(p)),
      batchUpdate: (p) => withRetry(() => base.spreadsheets.batchUpdate(p)),
    },
  };
}

function getRetryableDrive() {
  const base = getDrive();
  return {
    files: {
      create: (p) => withRetry(() => base.files.create(p)),
      get: (p) => withRetry(() => base.files.get(p)),
      update: (p) => withRetry(() => base.files.update(p)),
      delete: (p) => withRetry(() => base.files.delete(p)),
      list: (p) => withRetry(() => base.files.list(p)),
    },
    permissions: {
      create: (p) => withRetry(() => base.permissions.create(p)),
    },
  };
}

module.exports = { withRetry, getRetryableSheets, getRetryableDrive };