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
];

const getSpreadsheet = () => SpreadsheetApp.openById(SPREADSHEET_ID);

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

  if (params.sheet) {
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
