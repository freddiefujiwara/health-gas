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
// In-memory CacheService backing store
const cacheStore = new Map();
global.CacheService = {
  getScriptCache: vi.fn().mockImplementation(() => ({
    get: vi.fn().mockImplementation((key) => (cacheStore.has(key) ? cacheStore.get(key) : null)),
    getAll: vi.fn().mockImplementation((keys) => {
      const out = {};
      keys.forEach((key) => {
        if (cacheStore.has(key)) {
          out[key] = cacheStore.get(key);
        }
      });
      return out;
    }),
    putAll: vi.fn().mockImplementation((entries) => {
      Object.entries(entries).forEach(([k, v]) => cacheStore.set(k, v));
    }),
  })),
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
    cacheStore.clear();
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

  it("should return sheet_not_found when a known sheet is missing from the spreadsheet", () => {
    mockSpreadsheet.getSheetByName.mockReturnValue(null);
    const e = { parameter: { sheet: "steps" } };
    doGet(e);
    const content = mockContentService.setContent.mock.calls[0][0];
    expect(JSON.parse(content)).toEqual({
      error: "sheet_not_found",
      sheet: "steps",
    });
  });

  it("should format Date cells with Utilities.formatDate", () => {
    const cellDate = new Date("2026-08-23T00:00:00.000Z");
    mockRange.getValues.mockReturnValueOnce([
      ["date", "steps"],
      [cellDate, 6148],
    ]);
    const e = { parameter: { sheet: "steps" } };
    doGet(e);
    const content = mockContentService.setContent.mock.calls[0][0];
    expect(JSON.parse(content)).toEqual([
      { date: "2026-08-23T00:00:00+09:00", steps: 6148 },
    ]);
    expect(global.Utilities.formatDate).toHaveBeenCalledWith(
      cellDate,
      "Asia/Tokyo",
      "yyyy-MM-dd'T'HH:mm:ssXXX"
    );
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
    cacheStore.clear();
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

  const rejectWithPayload = (overrides) => {
    global.UrlFetchApp.fetch.mockImplementation(() => ({
      getResponseCode: vi.fn().mockReturnValue(200),
      getContentText: vi.fn().mockReturnValue(
        JSON.stringify({ ...tokenInfoPayload, ...overrides })
      ),
    }));
    doGet({ parameter: { id_token: "good-token" } });
    return lastJson();
  };

  it("rejects a token with an invalid issuer", () => {
    expect(rejectWithPayload({ iss: "https://example.com" })).toEqual({
      status: 401,
      error: "invalid iss",
    });
  });

  it("rejects a token with an invalid audience", () => {
    expect(rejectWithPayload({ aud: "other-client-id" })).toEqual({
      status: 401,
      error: "invalid aud",
    });
  });

  it("rejects an expired token", () => {
    const past = String(Math.floor(Date.now() / 1000) - 60);
    expect(rejectWithPayload({ exp: past })).toEqual({
      status: 401,
      error: "token expired",
    });
  });

  it("rejects a token without an expiry", () => {
    expect(rejectWithPayload({ exp: undefined })).toEqual({
      status: 401,
      error: "token expired",
    });
  });

  it("rejects a token with a non-numeric expiry", () => {
    expect(rejectWithPayload({ exp: "not-a-number" })).toEqual({
      status: 401,
      error: "token expired",
    });
  });

  it("rejects a token whose email is not verified", () => {
    expect(rejectWithPayload({ email_verified: "false" })).toEqual({
      status: 401,
      error: "email not verified",
    });
  });

  it("rejects a token without an email", () => {
    expect(rejectWithPayload({ email: "" })).toEqual({
      status: 401,
      error: "missing email",
    });
  });

  it("rejects when GOOGLE_OAUTH_CLIENT_ID is not configured", () => {
    delete scriptProperties.GOOGLE_OAUTH_CLIENT_ID;
    doGet({ parameter: { id_token: "good-token" } });
    expect(lastJson()).toEqual({
      status: 401,
      error: "missing GOOGLE_OAUTH_CLIENT_ID",
    });
    expect(global.UrlFetchApp.fetch).not.toHaveBeenCalled();
  });

  it("rejects when PropertiesService is unavailable", () => {
    const original = global.PropertiesService;
    global.PropertiesService = undefined;
    try {
      doGet({ parameter: { id_token: "good-token" } });
      expect(lastJson()).toEqual({
        status: 401,
        error: "missing GOOGLE_OAUTH_CLIENT_ID",
      });
    } finally {
      global.PropertiesService = original;
    }
  });

  it("rejects when PropertiesService has no script properties store", () => {
    const original = global.PropertiesService;
    global.PropertiesService = {};
    try {
      doGet({ parameter: { id_token: "good-token" } });
      expect(lastJson()).toEqual({
        status: 401,
        error: "missing GOOGLE_OAUTH_CLIENT_ID",
      });
    } finally {
      global.PropertiesService = original;
    }
  });

  it("accepts the bare issuer spelling and a boolean verified flag", () => {
    global.UrlFetchApp.fetch.mockImplementation(() => ({
      getResponseCode: vi.fn().mockReturnValue(200),
      getContentText: vi.fn().mockReturnValue(
        JSON.stringify({
          ...tokenInfoPayload,
          iss: "accounts.google.com",
          email_verified: true,
        })
      ),
    }));
    doGet({ parameter: { id_token: "good-token", sheet: "steps" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
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
    cacheStore.clear();
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

  it("serves data when the stored session dates are Date objects", () => {
    const token = "date-cells-token";
    const expiresAt = new Date(Date.now() + 1800_000);
    shareValues.push([
      shareHashOf(token),
      new Date(Date.now() - 1000),
      expiresAt,
      "",
      0,
      "",
    ]);
    doGet({ parameter: { share_token: token } });
    const data = lastJson();
    expect(data.steps).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(data._share.expires_at).toBe(expiresAt.toISOString());
    expect(shareValues[1][4]).toBe(1);
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

describe("doGet owner sessions", () => {
  const SESSION_HEADERS = [
    "token_hash",
    "issued_at",
    "expires_at",
    "revoked_at",
    "access_count",
    "last_accessed_at",
  ];

  let ownerValues = [];
  const mockOwnerSheet = {
    getDataRange: vi.fn().mockImplementation(() => ({
      getValues: vi.fn().mockImplementation(() => ownerValues),
    })),
    appendRow: vi.fn().mockImplementation((row) => {
      ownerValues.push([...row]);
    }),
    getRange: vi.fn().mockImplementation((row, col) => ({
      setValue: vi.fn().mockImplementation((value) => {
        ownerValues[row - 1][col - 1] = value;
      }),
    })),
  };

  beforeEach(() => {
    vi.clearAllMocks();
    cacheStore.clear();
    uuidCounter = 0;
    ownerValues = [[...SESSION_HEADERS]];
    shareValues = [[...SESSION_HEADERS]];
    mockSpreadsheet.insertSheet.mockReturnValue(mockOwnerSheet);
    mockSpreadsheet.getSheetByName.mockImplementation((name) => {
      if (name === "owner_sessions") return mockOwnerSheet;
      if (name === "share_sessions") return mockShareSheet;
      return mockSheet;
    });
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
  const exchange = () => {
    doGet({ parameter: { action: "exchange_session", id_token: "good-token" } });
    return lastJson().session_token;
  };

  it("requires an ID token to exchange for an owner session", () => {
    doGet({ parameter: { action: "exchange_session" } });
    expect(lastJson()).toEqual({ status: 401, error: "missing id token" });
  });

  it("issues a 30-day owner session and stores only its hash", () => {
    const token = exchange();
    const result = lastJson();
    expect(result.ttl_seconds).toBe(30 * 24 * 60 * 60);
    expect(token).toMatch(/^[0-9a-f]{64}$/);
    expect(ownerValues).toHaveLength(2);
    expect(ownerValues[1][0]).toBe(shareHashOf(token));
    expect(ownerValues[1][0]).not.toBe(token);
  });

  it("serves data with a valid owner session token", () => {
    const token = exchange();
    doGet({ parameter: { session_token: token, sheet: "steps" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(ownerValues[1][4]).toBe(1);
  });

  it("rejects an unknown owner session token", () => {
    doGet({ parameter: { session_token: "no-such-token" } });
    expect(lastJson()).toEqual({ status: 401, error: "session expired or invalid" });
  });

  it("keeps previous owner sessions valid when a new one is exchanged", () => {
    const first = exchange();
    const second = exchange();
    expect(first).not.toBe(second);
    doGet({ parameter: { session_token: first, sheet: "steps" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    doGet({ parameter: { session_token: second, sheet: "steps" } });
    expect(lastJson()).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });

  it("revokes the owner session via the revoke_session action", () => {
    const token = exchange();
    doGet({ parameter: { action: "revoke_session", session_token: token } });
    expect(lastJson()).toEqual({ revoked: true });
    doGet({ parameter: { session_token: token } });
    expect(lastJson()).toEqual({ status: 401, error: "session expired or invalid" });
  });

  it("accepts an owner session as proof for issuing a share link", () => {
    const token = exchange();
    doGet({ parameter: { action: "issue_share", session_token: token } });
    expect(lastJson().share_token).toMatch(/^[0-9a-f]{64}$/);
  });
});

describe("staging spreadsheet override", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    cacheStore.clear();
    mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);
  });

  it("opens the SPREADSHEET_ID script property when set", () => {
    scriptProperties.DEBUG = "true";
    scriptProperties.SPREADSHEET_ID = "staging-file-id";
    SpreadsheetApp.openById.mockClear();
    doGet({ parameter: { sheet: "steps" } });
    expect(SpreadsheetApp.openById).toHaveBeenCalledWith("staging-file-id");
    delete scriptProperties.SPREADSHEET_ID;
  });

  it("falls back to the built-in spreadsheet id when the property is unset", () => {
    scriptProperties.DEBUG = "true";
    SpreadsheetApp.openById.mockClear();
    doGet({ parameter: { sheet: "steps" } });
    expect(SpreadsheetApp.openById).toHaveBeenCalledWith(
      "1a80iQ8kgQ2mZaVPfxIfIpAn8ugyKkx0coNZa50vKR5A"
    );
  });
});

describe("sheet read cache", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    cacheStore.clear();
    mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);
    cacheStore.clear();
  });

  it("reads sheets once and serves the second full request from cache", () => {
    doGet({ parameter: {} });
    const callsAfterFirst = mockRange.getValues.mock.calls.length;
    expect(callsAfterFirst).toBeGreaterThan(0);
    doGet({ parameter: {} });
    expect(mockRange.getValues.mock.calls.length).toBe(callsAfterFirst);
  });

  it("serves a single-sheet request from the cache warmed by a full request", () => {
    doGet({ parameter: {} });
    const calls = mockRange.getValues.mock.calls.length;
    doGet({ parameter: { sheet: "steps" } });
    expect(mockRange.getValues.mock.calls.length).toBe(calls);
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    expect(JSON.parse(content)).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });
});
