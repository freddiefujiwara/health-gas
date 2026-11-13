const SPREADSHEET_ID = '1a80iQ8kgQ2mZaVPfxIfIpAn8ugyKkx0coNZa50vKR5A';

// Environment separation (staging): when the SPREADSHEET_ID script
// property is set, it wins over the built-in production ID above, so a
// standalone staging project can serve a dummy-data copy of the
// spreadsheet. When the property is unset, behavior is unchanged.

// Sheets exposed by the API. The first row of each sheet is used as keys.
const SHEET_NAMES = [
  "overview",
  "daily_nutrition",
  "steps",
  "meals",
  "period_summaries",
  "body_composition",
  "body_segments",
  "medication_plan",
  "medication_log",
  "goals_settings",
  "data_dictionary",
  "lab_results",
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

const issueSession_ = (sheetName, tokenKey) => {
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresIso = new Date(now.getTime() + SHARE_TTL_SECONDS * 1000).toISOString();
  revokeActiveSessions_(sheetName, nowIso);
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
// later calls. Issuing a new one revokes the previous ones, and the
// frontend revokes its session on logout.

const OWNER_SHEET_NAME = 'owner_sessions';

const issueOwnerSession_ = () => issueSession_(OWNER_SHEET_NAME, 'session_token');

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
    .filter(row => row.some(cell => cell !== '' && cell !== null))
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

export function doGet(e) {
  const params = (e && e.parameter) || {};

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
    if (!SHEET_NAMES.includes(params.sheet)) {
      return createJsonOutput({ error: 'sheet_not_found', sheet: params.sheet });
    }
    const sheet = getSheet(params.sheet);
    if (!sheet) {
      return createJsonOutput({ error: 'sheet_not_found', sheet: params.sheet });
    }
    return createJsonpOutput(
      sheetValuesToObjects(sheet.getDataRange().getValues()),
      params.callback
    );
  }

  const data = {};
  SHEET_NAMES.forEach(name => {
    const sheet = getSheet(name);
    if (sheet) {
      data[name] = sheetValuesToObjects(sheet.getDataRange().getValues());
    }
  });
  if (shareSession) {
    data._share = { expires_at: shareSession.expires_at };
  }
  return createJsonpOutput(data, params.callback);
}
