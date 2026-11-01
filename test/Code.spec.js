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
  DigestAlgorithm: { SHA_256: "SHA_256" },
  computeDigest: vi.fn().mockImplementation((algo, str) =>
    Array.from({ length: 32 }, (_, i) => String(str).charCodeAt(i % String(str).length) % 128)
  ),
  getUuid: vi.fn().mockImplementation(
    () => `00000000-0000-4000-8000-${String(++uuidCounter).padStart(12, "0")}`
  ),
};
let uuidCounter = 0;

const shareHashOf = (token) =>
  Array.from({ length: 32 }, (_, i) => token.charCodeAt(i % token.length) % 128)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");

let shareValues = [];
const mockShareSheet = {
  getDataRange: vi.fn().mockImplementation(() => ({
    getValues: vi.fn().mockImplementation(() => shareValues),
  })),
  appendRow: vi.fn().mockImplementation((row) => {
    shareValues.push([...row]);
  }),
  getRange: vi.fn().mockImplementation((row, col) => ({
    setValue: vi.fn().mockImplementation((value) => {
      shareValues[row - 1][col - 1] = value;
    }),
  })),
};
mockSpreadsheet.insertSheet = vi.fn().mockReturnValue(mockShareSheet);
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

describe("doGet share sessions", () => {
  const SHARE_HEADERS = [
    "token_hash",
    "issued_at",
    "expires_at",
    "revoked_at",
    "access_count",
    "last_accessed_at",
  ];

  beforeEach(() => {
    vi.clearAllMocks();
    uuidCounter = 0;
    shareValues = [[...SHARE_HEADERS]];
    mockSpreadsheet.getSheetByName.mockImplementation((name) =>
      name === "share_sessions" ? mockShareSheet : mockSheet
    );
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
  const issueToken = () => {
    doGet({ parameter: { action: "issue_share", id_token: "good-token" } });
    return lastJson().share_token;
  };

  it("requires an ID token to issue a share session", () => {
    doGet({ parameter: { action: "issue_share" } });
    expect(lastJson()).toEqual({ status: 401, error: "missing id token" });
  });

  it("issues a 30-day share session and stores only its hash", () => {
    const before = Date.now();
    doGet({ parameter: { action: "issue_share", id_token: "good-token" } });
    const result = lastJson();
    expect(result.ttl_seconds).toBe(30 * 24 * 60 * 60);
    expect(result.share_token).toMatch(/^[0-9a-f]{64}$/);
    expect(Date.parse(result.expires_at) - before).toBeGreaterThan(29 * 24 * 60 * 60 * 1000);
    expect(Date.parse(result.expires_at) - before).toBeLessThanOrEqual(30 * 24 * 60 * 60 * 1000 + 5000);
    expect(shareValues).toHaveLength(2);
    expect(shareValues[1][0]).toBe(shareHashOf(result.share_token));
    expect(shareValues[1][0]).not.toBe(result.share_token);
    expect(shareValues[1][4]).toBe(0);
  });

  it("creates the share sheet with headers when it does not exist yet", () => {
    shareValues = [];
    let shareSheetExists = false;
    mockSpreadsheet.insertSheet.mockImplementation(() => {
      shareSheetExists = true;
      return mockShareSheet;
    });
    mockSpreadsheet.getSheetByName.mockImplementation((name) =>
      name === "share_sessions" ? (shareSheetExists ? mockShareSheet : null) : mockSheet
    );
    doGet({ parameter: { action: "issue_share", id_token: "good-token" } });
    expect(mockSpreadsheet.insertSheet).toHaveBeenCalledWith("share_sessions");
    expect(shareValues[0]).toEqual(SHARE_HEADERS);
    expect(shareValues).toHaveLength(2);
  });

  it("revokes the previous session when a new one is issued", () => {
    const first = issueToken();
    const second = issueToken();
    expect(first).not.toBe(second);
    expect(shareValues[1][3]).not.toBe("");
    expect(shareValues[2][3]).toBe("");
    doGet({ parameter: { share_token: first } });
    expect(lastJson()).toEqual({ status: 401, error: "share session expired or invalid" });
  });

  it("serves data with a valid share token and counts the access", () => {
    const token = issueToken();
    doGet({ parameter: { share_token: token } });
    const data = lastJson();
    expect(data.steps).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(data._share.expires_at).toBe(shareValues[1][2]);
    expect(shareValues[1][4]).toBe(1);
    expect(shareValues[1][5]).not.toBe("");
  });

  it("rejects an expired share token", () => {
    const token = "expired-token";
    shareValues.push([
      shareHashOf(token),
      new Date(Date.now() - 3600_000).toISOString(),
      new Date(Date.now() - 1800_000).toISOString(),
      "",
      0,
      "",
    ]);
    doGet({ parameter: { share_token: token } });
    expect(lastJson()).toEqual({ status: 401, error: "share session expired or invalid" });
  });

  it("rejects a revoked share token", () => {
    const token = "revoked-token";
    shareValues.push([
      shareHashOf(token),
      new Date().toISOString(),
      new Date(Date.now() + 1800_000).toISOString(),
      new Date().toISOString(),
      0,
      "",
    ]);
    doGet({ parameter: { share_token: token } });
    expect(lastJson()).toEqual({ status: 401, error: "share session expired or invalid" });
  });

  it("rejects an unknown share token", () => {
    doGet({ parameter: { share_token: "no-such-token" } });
    expect(lastJson()).toEqual({ status: 401, error: "share session expired or invalid" });
  });

  it("revokes the active session via the revoke action", () => {
    const token = issueToken();
    doGet({ parameter: { action: "revoke_share", id_token: "good-token" } });
    expect(lastJson()).toEqual({ revoked: true });
    doGet({ parameter: { share_token: token } });
    expect(lastJson()).toEqual({ status: 401, error: "share session expired or invalid" });
  });
});
