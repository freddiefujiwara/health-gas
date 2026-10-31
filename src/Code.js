const SPREADSHEET_ID = '1a80iQ8kgQ2mZaVPfxIfIpAn8ugyKkx0coNZa50vKR5A';

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

const getSpreadsheet = () => SpreadsheetApp.openById(SPREADSHEET_ID);

// --- Authentication (same scheme as asset-gas) ---
// Script Properties required:
//   AVAILABLE_GMAILS: comma-separated emails allowed to read the data
//   GOOGLE_OAUTH_CLIENT_ID: OAuth client ID used to verify ID tokens
//   DEBUG: when "true", authentication is skipped (for local checks only)

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

const isDebugMode_ = () => getScriptProperty_('DEBUG') === 'true';

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

// Helper to create the final output for doGet
const createJsonpOutput = (data, callback) => {
  const output = ContentService.createTextOutput();
  if (callback) {
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

  try {
    if (!isDebugMode_()) {
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
  return createJsonpOutput(data, params.callback);
}
