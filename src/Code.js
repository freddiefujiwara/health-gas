const SPREADSHEET_ID = '1a80iQ8kgQ2mZaVPfxIfIpAn8ugyKkx0coNZa50vKR5A';

// Environment separation (staging): when the SPREADSHEET_ID script
// property is set, it wins over the built-in production ID above, so a
// standalone staging project can serve a dummy-data copy of the
// spreadsheet. When the property is unset, behavior is unchanged.

// Sheets exposed by the API. The first row of each sheet is used as keys.
// period_summaries is not a sheet read: it is computed from
// daily_nutrition on every request (see computePeriodSummaries_).
const SHEET_NAMES = [
  "overview",
  "daily_nutrition",
  "steps",
  "meals",
  "body_composition",
  "body_segments",
  "goals_settings",
  "data_dictionary",
  "lab_results",
  "chat_activity",
  "chat_activity_hours",
  "chat_activity_daily_hours",
  "sleep_estimates",
];

const getSpreadsheet = () =>
  SpreadsheetApp.openById(getScriptProperty_('SPREADSHEET_ID') || SPREADSHEET_ID);

// --- Temporary share sessions (doctor viewing links) ---
// A logged-in owner can issue a URL token that grants read-only access for
// SHARE_TTL_SECONDS without a Google account. Sessions are stored in a
// dedicated sheet (token hashes only, never raw tokens), which doubles as
// an audit log. This sheet is NOT part of SHEET_NAMES and is never exposed.

const SHARE_SHEET_NAME = 'share_sessions';
// One month, so the owner can hand the doctor a link well before a visit.
const SHARE_TTL_SECONDS = 30 * 24 * 60 * 60;
const SHARE_HEADERS = [
  'token_hash',
  'issued_at',
  'expires_at',
  'revoked_at',
  'access_count',
  'last_accessed_at',
];

const getSessionSheet_ = sheetName => {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(SHARE_HEADERS);
  }
  return sheet;
};

const hashShareToken_ = token =>
  Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token))
    .map(byte => ((byte % 256) + 256) % 256)
    .map(byte => byte.toString(16).padStart(2, '0'))
    .join('');

const readSessionRows_ = sheetName => {
  const sheet = getSessionSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  return { sheet, values };
};

const shareCellIso_ = value => {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
};

const revokeActiveSessions_ = (sheetName, nowIso) => {
  const { sheet, values } = readSessionRows_(sheetName);
  values.forEach((row, index) => {
    if (index === 0) {
      return;
    }
    const revokedAt = shareCellIso_(row[3]);
    const expiresAt = shareCellIso_(row[2]);
    if (!revokedAt && expiresAt && expiresAt > nowIso) {
      sheet.getRange(index + 1, 4).setValue(nowIso);
    }
  });
};

const revokeActiveShareSessions_ = nowIso =>
  revokeActiveSessions_(SHARE_SHEET_NAME, nowIso);

const issueSession_ = (sheetName, tokenKey, { revokePrevious = true } = {}) => {
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresIso = new Date(now.getTime() + SHARE_TTL_SECONDS * 1000).toISOString();
  if (revokePrevious) {
    revokeActiveSessions_(sheetName, nowIso);
  }
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  getSessionSheet_(sheetName).appendRow([hashShareToken_(token), nowIso, expiresIso, '', 0, '']);
  return {
    [tokenKey]: token,
    issued_at: nowIso,
    expires_at: expiresIso,
    ttl_seconds: SHARE_TTL_SECONDS,
  };
};

const issueShareSession_ = () => issueSession_(SHARE_SHEET_NAME, 'share_token');

const revokeShareSession_ = () => {
  revokeActiveShareSessions_(new Date().toISOString());
  return { revoked: true };
};

const validateSessionToken_ = (sheetName, token) => {
  if (!token) {
    return null;
  }
  const hash = hashShareToken_(token);
  const { sheet, values } = readSessionRows_(sheetName);
  const nowIso = new Date().toISOString();
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    if (row[0] !== hash) {
      continue;
    }
    if (shareCellIso_(row[3])) {
      return null;
    }
    const expiresAt = shareCellIso_(row[2]);
    if (!expiresAt || expiresAt <= nowIso) {
      return null;
    }
    sheet.getRange(index + 1, 5).setValue(Number(row[4] || 0) + 1);
    sheet.getRange(index + 1, 6).setValue(nowIso);
    return { expires_at: expiresAt };
  }
  return null;
};

const validateShareToken_ = token => validateSessionToken_(SHARE_SHEET_NAME, token);

// --- Owner sessions (long-lived sign-in) ---
// Google ID tokens expire after about an hour, which used to drop the
// owner back at the sign-in gate. After a Google sign-in the frontend
// exchanges the ID token for one of these 30-day owner sessions (same
// storage scheme as share sessions, separate sheet) and uses it for all
// later calls. Owner sessions coexist across devices: issuing a new one
// does NOT revoke the others (unlike share links, where reissuing must
// invalidate the old link). The frontend revokes its own session on
// logout.

const OWNER_SHEET_NAME = 'owner_sessions';

const issueOwnerSession_ = () =>
  issueSession_(OWNER_SHEET_NAME, 'session_token', { revokePrevious: false });

const validateOwnerToken_ = token => validateSessionToken_(OWNER_SHEET_NAME, token);

const revokeOwnerSession_ = token => {
  if (token) {
    const hash = hashShareToken_(token);
    const { sheet, values } = readSessionRows_(OWNER_SHEET_NAME);
    const nowIso = new Date().toISOString();
    values.forEach((row, index) => {
      if (index > 0 && row[0] === hash && !shareCellIso_(row[3])) {
        sheet.getRange(index + 1, 4).setValue(nowIso);
      }
    });
  }
  return { revoked: true };
};

// --- Authentication (same scheme as asset-gas) ---
// Script Properties required:
//   AVAILABLE_GMAILS: comma-separated emails allowed to read the data
//   GOOGLE_OAUTH_CLIENT_ID: OAuth client ID used to verify ID tokens

const getScriptProperties_ = () => {
  if (typeof PropertiesService === 'undefined' || !PropertiesService.getScriptProperties) {
    return null;
  }
  return PropertiesService.getScriptProperties();
};

const getScriptProperty_ = key => {
  const properties = getScriptProperties_();
  if (!properties || !properties.getProperty) {
    return '';
  }
  return properties.getProperty(key) || '';
};

const extractIdToken_ = e => {
  const headers = (e && e.headers) || {};
  const auth = headers.Authorization || headers.authorization || '';
  const bearerMatch = String(auth).match(/^Bearer\s+(.+)$/i);
  if (bearerMatch) {
    return bearerMatch[1];
  }
  return (e && e.parameter && e.parameter.id_token) || '';
};

const getAllowedEmails_ = () =>
  getScriptProperty_('AVAILABLE_GMAILS')
    .split(',')
    .map(value => value.trim().toLowerCase())
    .filter(Boolean);

const verifyGoogleIdTokenOrThrow_ = idToken => {
  if (!idToken) {
    throw new Error('missing id token');
  }
  const oauthClientId = getScriptProperty_('GOOGLE_OAUTH_CLIENT_ID');
  if (!oauthClientId) {
    throw new Error('missing GOOGLE_OAUTH_CLIENT_ID');
  }
  const response = UrlFetchApp.fetch(
    `https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(idToken)}`,
    { muteHttpExceptions: true }
  );
  if (response.getResponseCode() !== 200) {
    throw new Error('token verification failed');
  }
  const payload = JSON.parse(response.getContentText());
  const validIssuer =
    payload.iss === 'accounts.google.com' || payload.iss === 'https://accounts.google.com';
  if (!validIssuer) {
    throw new Error('invalid iss');
  }
  if (payload.aud !== oauthClientId) {
    throw new Error('invalid aud');
  }
  const expSeconds = Number(payload.exp || 0);
  const nowSeconds = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(expSeconds) || expSeconds <= nowSeconds) {
    throw new Error('token expired');
  }
  const emailVerified = payload.email_verified === true || payload.email_verified === 'true';
  if (!emailVerified) {
    throw new Error('email not verified');
  }
  const email = String(payload.email || '').toLowerCase();
  if (!email) {
    throw new Error('missing email');
  }
  if (!getAllowedEmails_().includes(email)) {
    throw new Error('forbidden email');
  }
  return { email, sub: payload.sub };
};


// Owner proof for owner-only actions: a valid Google ID token, or a
// live owner session token (the post-sign-in steady state).
const requireOwnerAuth_ = e => {
  const params = (e && e.parameter) || {};
  if (params.session_token && validateOwnerToken_(params.session_token)) {
    return;
  }
  verifyGoogleIdTokenOrThrow_(extractIdToken_(e));
};

const getSheet = name => getSpreadsheet().getSheetByName(name);

// Helper to convert 2D array from a sheet to an array of objects
const sheetValuesToObjects = values => {
  const headers = values.shift();
  return values
    .filter(row => Array.isArray(row) && row.some(cell => cell !== '' && cell !== null))
    .map(row =>
      headers.reduce((obj, header, index) => {
        const value = row[index];
        if (value instanceof Date) {
          obj[header] = Utilities.formatDate(value, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
        } else {
          obj[header] = value === '' || value === undefined ? null : value;
        }
        return obj;
      }, {})
    );
};

// --- Sheet read cache (CacheService) ---
// Reading every sheet from the spreadsheet takes several seconds, while
// the data changes only a few times a day. Converted rows are cached as
// JSON for a few minutes; values are chunked because CacheService caps a
// single value at 100KB. The cache is best-effort: any failure falls back
// to a direct sheet read.
const SHEET_CACHE_TTL_SECONDS = 300;
const SHEET_CACHE_CHUNK_SIZE = 90000;

const readSheetObjects_ = name => {
  const cacheKey = `sheetcache:${name}`;
  try {
    const cache = CacheService.getScriptCache();
    const count = Number(cache.get(`${cacheKey}:count`) || 0);
    if (count > 0) {
      const keys = [];
      for (let i = 0; i < count; i += 1) {
        keys.push(`${cacheKey}:${i}`);
      }
      const chunks = cache.getAll(keys);
      const joined = keys.map(key => chunks[key] || '').join('');
      if (joined) {
        return JSON.parse(joined);
      }
    }
  } catch (error) {
    // Cache unavailable or stale chunks: read the sheet directly.
  }
  const sheet = getSheet(name);
  if (!sheet) {
    return null;
  }
  const objects = sheetValuesToObjects(sheet.getDataRange().getValues());
  try {
    const json = JSON.stringify(objects);
    const entries = {};
    let count = 0;
    for (let i = 0; i < json.length; i += SHEET_CACHE_CHUNK_SIZE) {
      entries[`${cacheKey}:${count}`] = json.slice(i, i + SHEET_CACHE_CHUNK_SIZE);
      count += 1;
    }
    entries[`${cacheKey}:count`] = String(count);
    CacheService.getScriptCache().putAll(entries, SHEET_CACHE_TTL_SECONDS);
  } catch (error) {
    // Cache is best-effort.
  }
  return objects;
};

// A JSONP callback must be a plain JavaScript identifier path; anything
// else would be reflected verbatim into a JavaScript response (the same
// validation bucket-list-gas applies).
const isValidCallback_ = callback =>
  typeof callback === 'string' &&
  /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback);

// Helper to create the final output for doGet
const createJsonpOutput = (data, callback) => {
  const output = ContentService.createTextOutput();
  if (callback && isValidCallback_(callback)) {
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    output.setContent(`${callback}&&${callback}(${JSON.stringify(data)});`);
  } else {
    output.setMimeType(ContentService.MimeType.JSON);
    output.setContent(JSON.stringify(data));
  }
  return output;
};

const createJsonOutput = data => createJsonpOutput(data, null);

// --- period_summaries: computed from daily_nutrition -------------------
// Weekly (Monday-Sunday) aggregates over the completed daily_nutrition
// rows. A week with all 7 days completed is reported with its Monday and
// Sunday; a partial week is reported over its completed days only.
const formatDateStr_ = dt => {
  const pad = n => String(n).padStart(2, '0');
  return `${dt.getUTCFullYear()}-${pad(dt.getUTCMonth() + 1)}-${pad(dt.getUTCDate())}`;
};

// A YYYY-MM-DD key that names a day which actually exists (the Date
// constructor silently rolls 2026-02-30 over to March 2).
const isRealDateKey_ = key => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) {
    return false;
  }
  const [y, m, d] = key.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
};

const mondayOfWeek_ = dateStr => {
  const [y, m, d] = dateStr.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() - ((dt.getUTCDay() + 6) % 7));
  return formatDateStr_(dt);
};

const addDays_ = (dateStr, days) => {
  const [y, m, d] = dateStr.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() + days);
  return formatDateStr_(dt);
};

const computePeriodSummaries_ = dailyRows => {
  // Some rows were stored as Date cells and arrive as ISO datetimes
  // ("2026-10-31T00:00:00+09:00"); only the date prefix matters here.
  const completed = (dailyRows || [])
    .filter(row => row && isRealDateKey_(String(row.date || '').slice(0, 10))
      && row.completeness === '完了'
      && weeklyRssNum_(row.intake_kcal) !== null
      && weeklyRssNum_(row.expenditure_kcal) !== null
      && weeklyRssNum_(row.balance_kcal) !== null)
    .map(row => ({ ...row, date: String(row.date).slice(0, 10) }));
  // One row per date: a duplicated date keeps its last row (the same
  // rule weeklyRssByDate_ applies), so a date is never counted twice.
  const byDate = new Map();
  completed.forEach(row => byDate.set(row.date, row));
  const deduped = [...byDate.values()].sort((a, b) => (a.date < b.date ? -1 : 1));
  const byWeek = new Map();
  deduped.forEach(row => {
    const monday = mondayOfWeek_(row.date);
    if (!byWeek.has(monday)) {
      byWeek.set(monday, []);
    }
    byWeek.get(monday).push(row);
  });
  return [...byWeek.keys()].sort().map(monday => {
    const rows = byWeek.get(monday);
    const days = rows.length;
    const totalIntake = rows.reduce((sum, row) => sum + weeklyRssNum_(row.intake_kcal), 0);
    const totalExpenditure = rows.reduce((sum, row) => sum + weeklyRssNum_(row.expenditure_kcal), 0);
    const totalBalance = rows.reduce((sum, row) => sum + weeklyRssNum_(row.balance_kcal), 0);
    const full = days === 7;
    return {
      start_date: full ? monday : rows[0].date,
      end_date: full ? addDays_(monday, 6) : rows[days - 1].date,
      days,
      average_intake_kcal: Math.round(totalIntake / days),
      average_expenditure_kcal: Math.round(totalExpenditure / days),
      total_balance_kcal: totalBalance,
      average_balance_kcal: Math.round(totalBalance / days),
      completeness: full ? '完了' : '途中集計',
      source: 'daily_nutrition 集計',
      notes: full ? '月〜日の週集計' : '完了日のみで集計',
    };
  });
};

export function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.action === 'weekly_rss') {
    return serveWeeklyRss_(params);
  }

  let shareSession = null;
  try {
    if (params.action === 'exchange_session') {
      verifyGoogleIdTokenOrThrow_(extractIdToken_(e));
      return createJsonOutput(issueOwnerSession_());
    }
    if (params.action === 'revoke_session') {
      return createJsonOutput(revokeOwnerSession_(params.session_token || ''));
    }
    if (params.action === 'issue_share' || params.action === 'revoke_share') {
      requireOwnerAuth_(e);
      return createJsonOutput(
        params.action === 'issue_share' ? issueShareSession_() : revokeShareSession_()
      );
    }
    if (params.session_token) {
      if (!validateOwnerToken_(params.session_token)) {
        throw new Error('session expired or invalid');
      }
    } else if (params.share_token) {
      shareSession = validateShareToken_(params.share_token);
      if (!shareSession) {
        throw new Error('share session expired or invalid');
      }
    } else {
      verifyGoogleIdTokenOrThrow_(extractIdToken_(e));
    }
  } catch (error) {
    const message = (error && error.message) || 'unauthorized';
    const status = message === 'forbidden email' ? 403 : 401;
    return createJsonOutput({ status, error: message });
  }

  if (params.sheet) {
    if (params.sheet === 'period_summaries') {
      const daily = readSheetObjects_('daily_nutrition');
      return createJsonpOutput(computePeriodSummaries_(daily || []), params.callback);
    }
    if (!SHEET_NAMES.includes(params.sheet)) {
      return createJsonOutput({ error: 'sheet_not_found', sheet: params.sheet });
    }
    const objects = readSheetObjects_(params.sheet);
    if (!objects) {
      return createJsonOutput({ error: 'sheet_not_found', sheet: params.sheet });
    }
    return createJsonpOutput(objects, params.callback);
  }

  const data = {};
  SHEET_NAMES.forEach(name => {
    const objects = readSheetObjects_(name);
    if (objects) {
      data[name] = objects;
    }
  });
  data.period_summaries = computePeriodSummaries_(data.daily_nutrition || []);
  if (shareSession) {
    data._share = { expires_at: shareSession.expires_at };
  }
  return createJsonpOutput(data, params.callback);
}

// --- Weekly review RSS (action=weekly_rss) ------------------------------
// GET /exec?action=weekly_rss&token=<redacted> returns a weekly
// review of the health data as an RSS 2.0 feed for /reader,
// mirroring the /social and /asset dashboard feeds: one item per
// week, delivered Mondays at 06:20 (Asia/Tokyo), covering the
// previous Monday-Sunday week in full. The item is computed on
// the fly from the same sheet data the JSON API serves. The guid
// is the slot's Monday date, so re-fetching never duplicates an
// item. The items carry real weight figures, so the feed is
// token-gated by the WEEKLY_RSS_TOKEN script property and fails
// closed: while the property is unset, or the token mismatches,
// every request is refused.

const WEEKLY_RSS_TOKEN_PROPERTY = 'WEEKLY_RSS_TOKEN';
const WEEKLY_RSS_TIME = '06:20:00';
const WEEKLY_RSS_CHANNEL = {
  title: '健康ウィークリー',
  link: 'https://freddiefujiwara.com/health/',
  description: '健康データの週次レビュー（毎週月曜6:20）',
};
const WEEKLY_RSS_SHEETS = [
  'body_composition',
  'daily_nutrition',
  'steps',
  'chat_activity',
  'sleep_estimates',
  'lab_results',
  'goals_settings',
];
const WEEKLY_RSS_JST_OFFSET_MS = 9 * 60 * 60 * 1000;
const WEEKLY_RSS_SUMMARY_CHARS = 140;
const WEEKLY_RSS_WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];
const WEEKLY_RSS_DEFAULT_KCAL_PER_KG = 7700;

// The JST calendar date (YYYY-MM-DD) containing the instant. Date-key
// arithmetic below reuses addDays_ / mondayOfWeek_ from the
// period_summaries section: for plain YYYY-MM-DD keys they agree with
// the JST-based computation.
const weeklyRssJstDateKey_ = ms => formatDateStr_(new Date(ms + WEEKLY_RSS_JST_OFFSET_MS));

const weeklyRssWeekday_ = dateKey =>
  WEEKLY_RSS_WEEKDAYS[new Date(Date.parse(dateKey + 'T00:00:00Z')).getUTCDay()];

const weeklyRssMd_ = dateKey => {
  const parts = dateKey.split('-');
  return Number(parts[1]) + '/' + Number(parts[2]);
};

const weeklyRssDayLabel_ = dateKey =>
  weeklyRssMd_(dateKey) + '（' + weeklyRssWeekday_(dateKey) + '）';

// The current delivery slot: the most recent Monday 06:20 JST.
// The covered week is the Monday-Sunday week that ended the day
// before the slot Monday. Returns { mondayKey, atMs, weekStartKey,
// weekEndKey }.
export const weeklyRssSlot_ = nowMs => {
  const todayKey = weeklyRssJstDateKey_(nowMs);
  const thisMondayKey = mondayOfWeek_(todayKey);
  const thisMondayAtMs = Date.parse(thisMondayKey + 'T' + WEEKLY_RSS_TIME + '+09:00');
  const mondayKey =
    thisMondayAtMs <= nowMs ? thisMondayKey : addDays_(thisMondayKey, -7);
  return {
    mondayKey,
    atMs: Date.parse(mondayKey + 'T' + WEEKLY_RSS_TIME + '+09:00'),
    weekStartKey: addDays_(mondayKey, -7),
    weekEndKey: addDays_(mondayKey, -1),
  };
};

const escapeWeeklyXmlText_ = value => {
  let text;
  try {
    text = String(value === null || value === undefined ? '' : value);
  } catch (error) {
    text = '';
  }
  return text
    // XML 1.0 cannot carry lone surrogates or C0 controls other than
    // tab / LF / CR: lone surrogates become U+FFFD, controls are cut.
    .replace(/([\uD800-\uDBFF][\uDC00-\uDFFF])|[\uD800-\uDFFF]/g, (match, pair) => pair || '\uFFFD')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
};

const createWeeklyXmlResponse_ = xml => {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.XML);
  output.setContent(xml);
  return output;
};

// Builds the RSS 2.0 document from weekly review items.
export const buildWeeklyRss_ = items => {
  const lines = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<rss version="2.0">',
    '<channel>',
    '<title>' + escapeWeeklyXmlText_(WEEKLY_RSS_CHANNEL.title) + '</title>',
    '<link>' + escapeWeeklyXmlText_(WEEKLY_RSS_CHANNEL.link) + '</link>',
    '<description>' + escapeWeeklyXmlText_(WEEKLY_RSS_CHANNEL.description) + '</description>',
  ];
  (items || []).forEach(item => {
    if (!item || !item.guid || !item.pubDate) {
      return;
    }
    lines.push(
      '<item>',
      '<title>' + escapeWeeklyXmlText_(item.title || '') + '</title>',
      '<link>' + escapeWeeklyXmlText_(WEEKLY_RSS_CHANNEL.link) + '</link>',
      '<guid isPermaLink="false">' + escapeWeeklyXmlText_(item.guid) + '</guid>',
      '<pubDate>' + item.pubDate + '</pubDate>',
      '<description>' + escapeWeeklyXmlText_(item.descriptionText || '') + '</description>',
      '</item>',
    );
  });
  lines.push('</channel>', '</rss>');
  return lines.join('\n');
};

// Numbers and numeric strings only: a bare Number() would invent
// measurements out of booleans (true -> 1), single-element arrays
// ([90.5] -> 90.5), blank strings ('' -> 0) and Date cells (epoch ms).
const weeklyRssNum_ = value => {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
};

// Count-like fields (steps, minutes, tallies): a negative count is a
// deformed measurement, not a value.
const weeklyRssCount_ = value => {
  const parsed = weeklyRssNum_(value);
  return parsed !== null && parsed >= 0 ? parsed : null;
};

// Total text coercion for fields printed into the feed: strings pass,
// numbers and booleans stringify, anything else (including objects
// whose toString is not callable) reads as empty instead of throwing
// or printing "[object Object]".
const weeklyRssText_ = value => {
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  return '';
};

const weeklyRssDate10_ = value => weeklyRssText_(value).slice(0, 10);

const weeklyRssFmtInt_ = value => String(Math.round(value)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

const weeklyRssFmt1_ = value => (Math.round(value * 10) / 10).toFixed(1);

const weeklyRssFmtSignedInt_ = value =>
  (value > 0 ? '+' : value < 0 ? '-' : '±') + weeklyRssFmtInt_(Math.abs(value));

const weeklyRssFmtSigned1_ = value =>
  (value > 0 ? '+' : value < 0 ? '-' : '±') + weeklyRssFmt1_(Math.abs(value));

const weeklyRssFmtDuration_ = minutes => {
  const total = Math.round(minutes);
  const hours = Math.floor(total / 60);
  const rest = total % 60;
  return hours + '時間' + String(rest).padStart(2, '0') + '分';
};

const weeklyRssByDate_ = rows => {
  const map = {};
  (Array.isArray(rows) ? rows : []).forEach(row => {
    if (row && row.date) {
      map[weeklyRssDate10_(row.date)] = row;
    }
  });
  return map;
};

const weeklyRssWeekDates_ = slot => {
  const dates = [];
  for (let i = 0; i < 7; i += 1) {
    dates.push(addDays_(slot.weekStartKey, i));
  }
  return dates;
};

// Empirical kcal-per-kg factor, the same computation the front
// end runs: anchors are the first and latest body-composition
// measurements on/after the first completed nutrition date, and
// the factor is the cumulative balance between them divided by
// the weight lost. Null when it cannot be computed.
const weeklyRssEmpirical_ = (bodyRows, nutritionRows) => {
  const completed = (Array.isArray(nutritionRows) ? nutritionRows : [])
    .filter(row => row && row.completeness === '完了' && weeklyRssNum_(row.balance_kcal) !== null)
    .map(row => ({ date: weeklyRssDate10_(row.date), balanceKcal: weeklyRssNum_(row.balance_kcal) }))
    .filter(row => isRealDateKey_(row.date))
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  if (!completed.length) {
    return null;
  }
  const firstCompletedDate = completed[0].date;
  const anchors = (Array.isArray(bodyRows) ? bodyRows : [])
    .filter(row => row && weeklyRssNum_(row.weight_kg) !== null)
    .map(row => ({ date: weeklyRssDate10_(row.date), weightKg: weeklyRssNum_(row.weight_kg) }))
    .filter(row => isRealDateKey_(row.date) && row.date >= firstCompletedDate)
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  if (anchors.length < 2) {
    return null;
  }
  const first = anchors[0];
  const last = anchors[anchors.length - 1];
  let totalBalance = 0;
  completed.forEach(row => {
    if (row.date > first.date && row.date <= last.date) {
      totalBalance += row.balanceKcal;
    }
  });
  const lossKg = first.weightKg - last.weightKg;
  if (totalBalance >= 0 || lossKg <= 0) {
    return null;
  }
  const kcalPerKg = Math.round(-totalBalance / lossKg);
  if (!Number.isFinite(kcalPerKg) || kcalPerKg < 1) {
    // A factor that rounds to 0 kcal/kg is unusable: the forecasts
    // divide by it and would print Infinity.
    return null;
  }
  return {
    kcalPerKg,
    fromDate: first.date,
    toDate: last.date,
    totalBalanceKcal: totalBalance,
  };
};

// Builds the weekly review item for the current slot. `data`
// holds the sheet rows keyed by sheet name, in the same shape
// the JSON API serves.
export const buildWeeklyReviewItem_ = (data, nowMs) => {
  const slot = weeklyRssSlot_(nowMs);
  const weekDates = weeklyRssWeekDates_(slot);
  // Sheet values arrive as arrays from readSheetObjects_; anything
  // else reads as an empty sheet rather than throwing here.
  const sheetRows = name => (data && Array.isArray(data[name]) ? data[name] : []);
  const bodyRows = sheetRows('body_composition')
    .filter(row => row && weeklyRssNum_(row.weight_kg) !== null)
    .map(row => ({ ...row, date: weeklyRssDate10_(row.date) }))
    .filter(row => isRealDateKey_(row.date))
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  const nutritionByDate = weeklyRssByDate_(sheetRows('daily_nutrition'));
  const stepsByDate = weeklyRssByDate_(sheetRows('steps'));
  const activityByDate = weeklyRssByDate_(sheetRows('chat_activity'));
  const sleepByDate = weeklyRssByDate_(sheetRows('sleep_estimates'));
  const goalByItem = {};
  sheetRows('goals_settings').forEach(row => {
    const item = weeklyRssText_(row && row.item);
    if (item) {
      goalByItem[item] = row;
    }
  });
  const goalValue = item => {
    const row = goalByItem[item];
    return row ? weeklyRssNum_(row.value) : null;
  };

  const lines = [];
  let predictedNowInfo = null;
  const section = title => {
    lines.push('');
    lines.push('■ ' + title);
  };

  // --- Body composition -------------------------------------------
  section('体重・体組成');
  const weekBody = bodyRows.filter(row => row.date >= slot.weekStartKey && row.date <= slot.weekEndKey);
  if (weekBody.length) {
    weekBody.forEach(row => {
      const parts = ['体重 ' + weeklyRssFmt1_(row.weight_kg) + ' kg'];
      if (weeklyRssNum_(row.body_fat_percent) !== null) {
        parts.push('体脂肪率 ' + weeklyRssFmt1_(row.body_fat_percent) + '%');
      }
      if (weeklyRssNum_(row.fat_mass_kg) !== null) {
        parts.push('脂肪量 ' + weeklyRssFmt1_(row.fat_mass_kg) + ' kg');
      }
      if (weeklyRssNum_(row.muscle_mass_kg) !== null) {
        parts.push('筋肉量 ' + weeklyRssFmt1_(row.muscle_mass_kg) + ' kg');
      }
      if (weeklyRssNum_(row.visceral_fat_level) !== null) {
        parts.push('内臓脂肪レベル ' + weeklyRssNum_(row.visceral_fat_level));
      }
      if (weeklyRssNum_(row.bmi) !== null) {
        parts.push('BMI ' + weeklyRssFmt1_(row.bmi));
      }
      lines.push(weeklyRssDayLabel_(row.date) + ' ' + parts.join('・'));
    });
  } else {
    lines.push('この週の実測はありません。');
  }
  const bodyUpToWeek = bodyRows.filter(row => row.date <= slot.weekEndKey);
  const latest = bodyUpToWeek.length ? bodyUpToWeek[bodyUpToWeek.length - 1] : null;
  const previous = bodyUpToWeek.length >= 2 ? bodyUpToWeek[bodyUpToWeek.length - 2] : null;
  if (latest) {
    let latestLine =
      '最新の実測：' + weeklyRssMd_(latest.date) + ' ' + weeklyRssFmt1_(latest.weight_kg) + ' kg';
    if (previous) {
      latestLine +=
        '（前回 ' + weeklyRssMd_(previous.date) + ' 比 ' +
        weeklyRssFmtSigned1_(latest.weight_kg - previous.weight_kg) + ' kg）';
    }
    lines.push(latestLine);
    // The starting reference comes from goals_settings (開始時体重,
    // whose unit_or_date carries the start date); the earliest
    // full-date measurement is the fallback. Partial dates such
    // as "2018-09" (month precision) never reach this point
    // because bodyRows only keeps full dates.
    const startGoalRow = goalByItem['開始時体重'];
    let startRef = null;
    if (startGoalRow && weeklyRssNum_(startGoalRow.value) !== null) {
      const dateMatch = weeklyRssText_(startGoalRow.unit_or_date).match(/\d{4}-\d{2}-\d{2}/);
      startRef = {
        date: dateMatch ? dateMatch[0] : null,
        weightKg: weeklyRssNum_(startGoalRow.value),
      };
    } else if (bodyRows.length) {
      startRef = { date: bodyRows[0].date, weightKg: weeklyRssNum_(bodyRows[0].weight_kg) };
    }
    if (startRef && startRef.weightKg > 0 && startRef.date !== latest.date) {
      const totalChange = latest.weight_kg - startRef.weightKg;
      const pct = (totalChange / startRef.weightKg) * 100;
      // A start weight near zero makes the percentage overflow; the
      // line is only meaningful with a finite percentage.
      if (Number.isFinite(pct)) {
        lines.push(
          '開始時（' + (startRef.date ? weeklyRssMd_(startRef.date) + ' ' : '') +
            weeklyRssFmt1_(startRef.weightKg) +
            ' kg）から累計 ' + weeklyRssFmtSigned1_(totalChange) + ' kg（' +
            weeklyRssFmtSigned1_(pct) + '%）',
        );
      }
    }
    const interimGoal = goalValue('中間目標');
    const finalGoal = goalValue('最終目標');
    const goalParts = [];
    if (interimGoal !== null) {
      goalParts.push(
        '中間目標 ' + weeklyRssFmt1_(interimGoal) + ' kg まで あと ' +
          weeklyRssFmt1_(Math.max(latest.weight_kg - interimGoal, 0)) + ' kg',
      );
    }
    if (finalGoal !== null) {
      goalParts.push(
        '最終目標 ' + weeklyRssFmt1_(finalGoal) + ' kg まで あと ' +
          weeklyRssFmt1_(Math.max(latest.weight_kg - finalGoal, 0)) + ' kg',
      );
    }
    if (goalParts.length) {
      lines.push(goalParts.join('／'));
    }
  }

  // Empirical factor, recent pace and forecasts (same figures
  // the overview cards show).
  const empirical = weeklyRssEmpirical_(sheetRows('body_composition'), sheetRows('daily_nutrition'));
  const kcalPerKg = empirical ? empirical.kcalPerKg : WEEKLY_RSS_DEFAULT_KCAL_PER_KG;
  if (empirical) {
    lines.push(
      '実測換算：体重 1 kg あたり ' + weeklyRssFmtInt_(empirical.kcalPerKg) + ' kcal（' +
        weeklyRssMd_(empirical.fromDate) + '〜' + weeklyRssMd_(empirical.toDate) +
        ' の累積収支 ' + weeklyRssFmtSignedInt_(empirical.totalBalanceKcal) + ' kcal から算出）',
    );
  } else {
    lines.push('実測換算はデータ不足のため、理論値 7,700 kcal/kg で計算しています。');
  }
  const completedUpToWeek = sheetRows('daily_nutrition')
    .filter(row => row && row.completeness === '完了'
      && isRealDateKey_(weeklyRssDate10_(row.date))
      && weeklyRssDate10_(row.date) <= slot.weekEndKey
      // A completed row whose balance does not parse is excluded
      // rather than counted as 0 (weeklyRssEmpirical_'s rule).
      && weeklyRssNum_(row.balance_kcal) !== null)
    .sort((a, b) => (weeklyRssDate10_(a.date) < weeklyRssDate10_(b.date) ? -1 : 1));
  const last30 = completedUpToWeek.slice(-30);
  const avgBalance30 = last30.length
    ? last30.reduce((acc, row) => acc + (weeklyRssNum_(row.balance_kcal) || 0), 0) / last30.length
    : null;
  if (avgBalance30 !== null) {
    lines.push('直近30日（完了日）の平均収支：' + weeklyRssFmtSignedInt_(Math.round(avgBalance30)) + ' kcal/日');
  }
  // Predicted weight at output time: the latest measurement plus
  // the completed balances recorded after it, converted with the
  // same factor (the end point of the front-end's actual-balance
  // projection).
  if (latest) {
    let sumAfterAnchor = 0;
    let lastAfterAnchorDate = null;
    completedUpToWeek.forEach(row => {
      const key = weeklyRssDate10_(row.date);
      if (key > latest.date) {
        sumAfterAnchor += weeklyRssNum_(row.balance_kcal) || 0;
        lastAfterAnchorDate = key;
      }
    });
    if (lastAfterAnchorDate) {
      predictedNowInfo = {
        weightKg: latest.weight_kg + sumAfterAnchor / kcalPerKg,
        asOf: lastAfterAnchorDate,
        sumKcal: sumAfterAnchor,
      };
      lines.push(
        '出力時点の推定体重：' + weeklyRssFmt1_(predictedNowInfo.weightKg) + ' kg（' +
          weeklyRssMd_(lastAfterAnchorDate) + ' 時点。最新実測 ' + weeklyRssMd_(latest.date) +
          ' の ' + weeklyRssFmt1_(latest.weight_kg) + ' kg から、完了した収支の累計 ' +
          weeklyRssFmtSignedInt_(sumAfterAnchor) + ' kcal を実測換算で換算）',
      );
    }
  }
  if (latest && avgBalance30 !== null) {
    const predicted30 = latest.weight_kg + (avgBalance30 * 30) / kcalPerKg;
    lines.push('このペースでの30日後予測：' + weeklyRssFmt1_(predicted30) + ' kg');
    if (avgBalance30 < 0) {
      const reachLine = (label, goalKg) => {
        if (goalKg === null) {
          return null;
        }
        const remaining = latest.weight_kg - goalKg;
        if (remaining <= 0) {
          return label + ' ' + weeklyRssFmt1_(goalKg) + ' kg には到達済みです';
        }
        const days = Math.ceil((remaining * kcalPerKg) / -avgBalance30);
        const reachDate = addDays_(latest.date, days);
        if (!isRealDateKey_(reachDate)) {
          // A near-zero daily pace pushes the forecast beyond the
          // representable date range; omit it rather than print NaN.
          return null;
        }
        return (
          label + ' ' + weeklyRssFmt1_(goalKg) + ' kg の到達予測：' +
          reachDate.replace(/-/g, '/')
        );
      };
      const interimLine = reachLine('中間目標', goalValue('中間目標'));
      const finalLine = reachLine('最終目標', goalValue('最終目標'));
      if (interimLine) {
        lines.push(interimLine);
      }
      if (finalLine) {
        lines.push(finalLine);
      }
    }
  }

  // --- Nutrition ---------------------------------------------------
  section('栄養（からサポ）');
  let weekIntake = 0;
  let weekExpenditure = 0;
  let weekBalance = 0;
  let completedDays = 0;
  weekDates.forEach(dateKey => {
    const row = nutritionByDate[dateKey];
    if (!row || weeklyRssNum_(row.intake_kcal) === null) {
      lines.push(weeklyRssDayLabel_(dateKey) + ' 記録なし');
      return;
    }
    const intake = weeklyRssNum_(row.intake_kcal);
    const expenditure = weeklyRssNum_(row.expenditure_kcal);
    const balance = weeklyRssNum_(row.balance_kcal);
    let line =
      weeklyRssDayLabel_(dateKey) +
      ' 摂取 ' + weeklyRssFmtInt_(intake) +
      '／消費 ' + (expenditure === null ? '—' : weeklyRssFmtInt_(expenditure)) +
      '／収支 ' + (balance === null ? '—' : weeklyRssFmtSignedInt_(balance)) + ' kcal';
    const completeness = weeklyRssText_(row.completeness);
    if (completeness && completeness !== '完了') {
      line += '（' + completeness + '）';
    }
    lines.push(line);
    if (weeklyRssNum_(row.protein_g) !== null) {
      const nutrientParts = ['タンパク質 ' + weeklyRssFmtInt_(row.protein_g) + ' g'];
      if (weeklyRssNum_(row.fat_g) !== null) {
        nutrientParts.push('脂質 ' + weeklyRssFmtInt_(row.fat_g) + ' g');
      }
      if (weeklyRssNum_(row.carbohydrates_g) !== null) {
        nutrientParts.push('炭水化物 ' + weeklyRssFmtInt_(row.carbohydrates_g) + ' g');
      }
      if (weeklyRssNum_(row.salt_g) !== null) {
        nutrientParts.push('塩分 ' + weeklyRssFmt1_(row.salt_g) + ' g');
      }
      lines.push('　栄養素：' + nutrientParts.join('・'));
    }
    if (row.completeness === '完了') {
      completedDays += 1;
      weekIntake += intake;
      weekExpenditure += expenditure || 0;
      weekBalance += balance || 0;
    }
  });
  if (completedDays) {
    lines.push(
      '週の合計（完了 ' + completedDays + ' 日）：摂取 ' + weeklyRssFmtInt_(weekIntake) +
        '／消費 ' + weeklyRssFmtInt_(weekExpenditure) +
        '／収支 ' + weeklyRssFmtSignedInt_(weekBalance) + ' kcal',
    );
    lines.push(
      '1日平均：摂取 ' + weeklyRssFmtInt_(weekIntake / completedDays) +
        '／消費 ' + weeklyRssFmtInt_(weekExpenditure / completedDays) +
        '／収支 ' + weeklyRssFmtSignedInt_(Math.round(weekBalance / completedDays)) + ' kcal',
    );
  }
  const energyGoal = goalValue('エネルギー目標');
  const proteinGoal = goalValue('タンパク質目標');
  const saltGoal = goalValue('塩分目標');
  if (energyGoal !== null || proteinGoal !== null || saltGoal !== null) {
    const targetParts = [];
    if (energyGoal !== null) {
      targetParts.push('エネルギー ' + weeklyRssFmtInt_(energyGoal) + ' kcal/日');
    }
    if (proteinGoal !== null) {
      targetParts.push('タンパク質 ' + weeklyRssFmtInt_(proteinGoal) + ' g/日');
    }
    if (saltGoal !== null) {
      targetParts.push('塩分 ' + weeklyRssFmt1_(saltGoal) + ' g/日');
    }
    lines.push('栄養指導の目標：' + targetParts.join('・'));
  }

  // --- Steps --------------------------------------------------------
  section('歩数');
  let totalSteps = 0;
  let stepsDays = 0;
  weekDates.forEach(dateKey => {
    const row = stepsByDate[dateKey];
    if (!row || weeklyRssCount_(row.steps) === null) {
      lines.push(weeklyRssDayLabel_(dateKey) + ' 記録なし');
      return;
    }
    const stepCount = weeklyRssCount_(row.steps);
    let line = weeklyRssDayLabel_(dateKey) + ' ' + weeklyRssFmtInt_(stepCount) + ' 歩';
    const extras = [];
    if (weeklyRssCount_(row.distance_walking_running_m) !== null) {
      extras.push(weeklyRssFmtInt_(weeklyRssCount_(row.distance_walking_running_m)) + ' m');
    }
    if (weeklyRssCount_(row.flights_climbed) !== null) {
      extras.push(weeklyRssFmtInt_(weeklyRssCount_(row.flights_climbed)) + ' 階');
    }
    if (extras.length) {
      line += '（' + extras.join('・') + '）';
    }
    const completeness = weeklyRssText_(row.completeness);
    if (completeness && completeness !== '完了') {
      line += '（' + completeness + '）';
    }
    lines.push(line);
    totalSteps += stepCount;
    stepsDays += 1;
  });
  if (stepsDays) {
    lines.push(
      '週の合計：' + weeklyRssFmtInt_(totalSteps) + ' 歩（1日平均 ' +
        weeklyRssFmtInt_(totalSteps / stepsDays) + ' 歩）',
    );
  }

  // --- Activity ------------------------------------------------------
  section('活動時間');
  let totalEngaged = 0;
  let activityDays = 0;
  weekDates.forEach(dateKey => {
    const row = activityByDate[dateKey];
    if (!row || weeklyRssCount_(row.engaged_minutes_est) === null) {
      lines.push(weeklyRssDayLabel_(dateKey) + ' 記録なし');
      return;
    }
    const engaged = weeklyRssCount_(row.engaged_minutes_est);
    let line = weeklyRssDayLabel_(dateKey) + ' ' + weeklyRssFmtDuration_(engaged);
    const bits = [];
    if (weeklyRssCount_(row.user_message_count) !== null) {
      bits.push('活動 ' + weeklyRssFmtInt_(weeklyRssCount_(row.user_message_count)) + ' 回');
    }
    if (weeklyRssCount_(row.late_night_message_count) !== null) {
      bits.push('深夜帯 ' + weeklyRssFmtInt_(weeklyRssCount_(row.late_night_message_count)) + ' 回');
    }
    if (bits.length) {
      line += '（' + bits.join('・') + '）';
    }
    lines.push(line);
    totalEngaged += engaged;
    activityDays += 1;
  });
  if (activityDays) {
    lines.push(
      '週の合計：' + weeklyRssFmtDuration_(totalEngaged) + '（1日平均 ' +
        weeklyRssFmtDuration_(totalEngaged / activityDays) + '）',
    );
  }

  // --- Sleep estimates ------------------------------------------------
  section('睡眠（推定）');
  let totalSleep = 0;
  let sleepDays = 0;
  weekDates.forEach(dateKey => {
    const row = sleepByDate[dateKey];
    if (!row || weeklyRssCount_(row.duration_minutes_est) === null) {
      lines.push(weeklyRssDayLabel_(dateKey) + ' 記録なし');
      return;
    }
    const duration = weeklyRssCount_(row.duration_minutes_est);
    lines.push(
      weeklyRssDayLabel_(dateKey) + ' ' + (weeklyRssText_(row.sleep_start_est) || '—') + '〜' +
        (weeklyRssText_(row.sleep_end_est) || '—') + '（' + weeklyRssFmtDuration_(duration) + '）',
    );
    totalSleep += duration;
    sleepDays += 1;
  });
  if (sleepDays) {
    lines.push('1日平均：' + weeklyRssFmtDuration_(totalSleep / sleepDays));
  }
  lines.push('※睡眠は活動記録からの推定です。');

  // --- Lab results ------------------------------------------------------
  section('検査結果');
  const weekLabs = sheetRows('lab_results').filter(row => {
    const date = weeklyRssDate10_(row && row.date);
    return isRealDateKey_(date) && date >= slot.weekStartKey && date <= slot.weekEndKey;
  });
  if (weekLabs.length) {
    const labDates = [...new Set(weekLabs.map(row => weeklyRssDate10_(row.date)))].sort();
    labDates.forEach(dateKey => {
      lines.push(weeklyRssDayLabel_(dateKey) + ' の検査');
      weekLabs
        .filter(row => weeklyRssDate10_(row.date) === dateKey)
        .forEach(row => {
          let line = '・' + weeklyRssText_(row.item) + ' ' + weeklyRssText_(row.value);
          const unit = weeklyRssText_(row.unit);
          if (unit) {
            line += ' ' + unit;
          }
          const flag = weeklyRssText_(row.flag);
          if (flag) {
            line += '（' + flag + '）';
          }
          lines.push(line);
        });
    });
  } else {
    lines.push('この週の新しい検査結果はありません。');
  }

  // --- Summary ----------------------------------------------------------
  let summary =
    weeklyRssMd_(slot.weekStartKey) + '〜' + weeklyRssMd_(slot.weekEndKey) + 'の週次レビュー。';
  if (latest) {
    summary +=
      '体重の最新実測は ' + weeklyRssFmt1_(latest.weight_kg) + ' kg（' +
      weeklyRssMd_(latest.date) + '）。';
  }
  if (predictedNowInfo) {
    summary += '推定体重は ' + weeklyRssFmt1_(predictedNowInfo.weightKg) + ' kg。';
  }
  if (completedDays) {
    summary +=
      '週の収支は合計 ' + weeklyRssFmtSignedInt_(weekBalance) + ' kcal（1日平均 ' +
      weeklyRssFmtSignedInt_(Math.round(weekBalance / completedDays)) + ' kcal）。';
  }
  if (stepsDays) {
    summary += '歩数は合計 ' + weeklyRssFmtInt_(totalSteps) + ' 歩。';
  }
  if (sleepDays) {
    summary += '推定睡眠は1日平均 ' + weeklyRssFmtDuration_(totalSleep / sleepDays) + '。';
  }
  summary = summary.slice(0, WEEKLY_RSS_SUMMARY_CHARS);

  return {
    title:
      weeklyRssMd_(slot.weekStartKey) + '〜' + weeklyRssMd_(slot.weekEndKey) + ' 健康ウィークリー',
    guid: 'health-weekly-' + slot.mondayKey,
    pubDate: new Date(slot.atMs).toUTCString(),
    descriptionText:
      summary + '\n-------\n' + lines.join('\n') + '\n詳しくは ' + WEEKLY_RSS_CHANNEL.link,
  };
};

// Serves action=weekly_rss. The token check comes first and
// fails closed: an unset WEEKLY_RSS_TOKEN refuses every request.
const serveWeeklyRss_ = params => {
  const expected = getScriptProperty_(WEEKLY_RSS_TOKEN_PROPERTY);
  if (!expected || !params.token || params.token !== expected) {
    return createJsonOutput({ ok: false, status: 403, error: 'invalid token' });
  }
  try {
    const data = {};
    WEEKLY_RSS_SHEETS.forEach(name => {
      data[name] = readSheetObjects_(name) || [];
    });
    const item = buildWeeklyReviewItem_(data, Date.now());
    return createWeeklyXmlResponse_(buildWeeklyRss_(item ? [item] : []));
  } catch (err) {
    return createJsonOutput({
      ok: false,
      status: 502,
      error: 'feed unavailable: ' + (err && err.message ? err.message : String(err)),
    });
  }
};
