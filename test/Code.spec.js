import { beforeEach, describe, expect, it, vi } from "vitest";
import { doGet } from "../src/Code.js";

// Mockup of Google Apps Script global objects
const mockRange = {
  getValues: vi.fn().mockImplementation(() => [
    ["date", "steps"],
    ["2026-08-23", 6148],
  ]),
};

const mockSheet = {
  getDataRange: vi.fn().mockReturnValue(mockRange),
};

const mockSpreadsheet = {
  getSheetByName: vi.fn().mockReturnValue(mockSheet),
};

const mockContentService = {
  createTextOutput: vi.fn().mockReturnThis(),
  setMimeType: vi.fn().mockReturnThis(),
  setContent: vi.fn(),
  MimeType: {
    JSON: "JSON",
    JAVASCRIPT: "JAVASCRIPT",
  },
};

// Assign mocks to global scope
global.SpreadsheetApp = {
  openById: vi.fn().mockReturnValue(mockSpreadsheet),
};
global.ContentService = mockContentService;
global.Utilities = {
  formatDate: vi.fn().mockReturnValue("2026-08-23T00:00:00+09:00"),
};
global.Logger = {
  log: vi.fn(),
};

// Script Properties: DEBUG=true by default so the data tests run without auth,
// mirroring how asset-gas is operated for checks.
const scriptProperties = { DEBUG: "true" };
global.PropertiesService = {
  getScriptProperties: vi.fn().mockImplementation(() => ({
    getProperty: vi.fn().mockImplementation((key) => scriptProperties[key] ?? null),
  })),
};

const tokenInfoPayload = {
  iss: "https://accounts.google.com",
  aud: "test-client-id",
  exp: String(Math.floor(Date.now() / 1000) + 3600),
  email_verified: "true",
  email: "zegenvs@gmail.com",
  sub: "1234567890",
};
global.UrlFetchApp = {
  fetch: vi.fn().mockImplementation(() => ({
    getResponseCode: vi.fn().mockReturnValue(200),
    getContentText: vi.fn().mockReturnValue(JSON.stringify(tokenInfoPayload)),
  })),
};

describe("doGet", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);
  });

  it("should return all sheets as JSON when no parameter is provided", () => {
    const e = { parameter: {} };
    doGet(e);
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JSON");
    const content = mockContentService.setContent.mock.calls[0][0];
    const data = JSON.parse(content);
    expect(data.steps).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(data.daily_nutrition).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });

  it("should return a single sheet when sheet parameter is provided", () => {
    const e = { parameter: { sheet: "steps" } };
    doGet(e);
    const content = mockContentService.setContent.mock.calls[0][0];
    expect(JSON.parse(content)).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });

  it("should return an error when the sheet does not exist", () => {
    mockSpreadsheet.getSheetByName.mockReturnValue(null);
    const e = { parameter: { sheet: "unknown" } };
    doGet(e);
    const content = mockContentService.setContent.mock.calls[0][0];
    expect(JSON.parse(content)).toEqual({
      error: "sheet_not_found",
      sheet: "unknown",
    });
  });

  it("should return JSONP when a callback is provided", () => {
    const e = { parameter: { sheet: "steps", callback: "myCallback" } };
    doGet(e);
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JAVASCRIPT");
    expect(mockContentService.setContent).toHaveBeenCalledWith(
      "myCallback&&myCallback(" +
        JSON.stringify([{ date: "2026-08-23", steps: 6148 }]) +
        ");"
    );
  });
});

describe("doGet authentication", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);
    for (const key of Object.keys(scriptProperties)) {
      delete scriptProperties[key];
    }
    Object.assign(scriptProperties, {
      AVAILABLE_GMAILS: "zegenvs@gmail.com",
      GOOGLE_OAUTH_CLIENT_ID: "test-client-id",
    });
    global.UrlFetchApp.fetch.mockImplementation(() => ({
      getResponseCode: vi.fn().mockReturnValue(200),
      getContentText: vi.fn().mockReturnValue(JSON.stringify(tokenInfoPayload)),
    }));
  });

  const lastJson = () => JSON.parse(mockContentService.setContent.mock.calls.at(-1)[0]);

  it("rejects requests without an ID token", () => {
    doGet({ parameter: {} });
    expect(lastJson()).toEqual({ status: 401, error: "missing id token" });
  });

  it("accepts a valid ID token for an allowed email", () => {
    doGet({ parameter: { id_token: "good-token", sheet: "steps" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(global.UrlFetchApp.fetch).toHaveBeenCalledWith(
      "https://oauth2.googleapis.com/tokeninfo?id_token=good-token",
      { muteHttpExceptions: true }
    );
  });

  it("accepts a Bearer token from the Authorization header", () => {
    doGet({ parameter: { sheet: "steps" }, headers: { Authorization: "Bearer header-token" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });

  it("rejects a token for a non-allowed email with 403", () => {
    global.UrlFetchApp.fetch.mockImplementation(() => ({
      getResponseCode: vi.fn().mockReturnValue(200),
      getContentText: vi.fn().mockReturnValue(
        JSON.stringify({ ...tokenInfoPayload, email: "other@example.com" })
      ),
    }));
    doGet({ parameter: { id_token: "good-token" } });
    expect(lastJson()).toEqual({ status: 403, error: "forbidden email" });
  });

  it("rejects when token verification fails at Google", () => {
    global.UrlFetchApp.fetch.mockImplementation(() => ({
      getResponseCode: vi.fn().mockReturnValue(400),
      getContentText: vi.fn().mockReturnValue("{}"),
    }));
    doGet({ parameter: { id_token: "bad-token" } });
    expect(lastJson()).toEqual({ status: 401, error: "token verification failed" });
  });

  it("skips authentication when DEBUG is true", () => {
    scriptProperties.DEBUG = "true";
    doGet({ parameter: { sheet: "steps" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(global.UrlFetchApp.fetch).not.toHaveBeenCalled();
  });
});
