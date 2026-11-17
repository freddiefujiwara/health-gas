import { beforeEach, describe, expect, it, vi } from "vitest";
import { buildWeeklyReviewItem_, buildWeeklyRss_, doGet, weeklyRssSlot_ } from "../src/Code.js";

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
    XML: "XML",
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

// Script Properties: authentication is always on, so the allowed account is
// configured by default and the data tests authenticate with the mocked ID
// token (see authedEvent below).
const scriptProperties = {
  AVAILABLE_GMAILS: "zegenvs@gmail.com",
  GOOGLE_OAUTH_CLIENT_ID: "test-client-id",
};
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

// Event carrying the mocked valid ID token, for tests that exercise the
// data endpoints rather than the authentication itself.
const authedEvent = (parameter = {}) => ({
  parameter: { ...parameter, id_token: "good-token" },
});

describe("doGet", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    cacheStore.clear();
    mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);
  });

  it("should return all sheets as JSON when no parameter is provided", () => {
    doGet(authedEvent());
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JSON");
    const content = mockContentService.setContent.mock.calls[0][0];
    const data = JSON.parse(content);
    expect(data.steps).toEqual([{ date: "2026-08-23", steps: 6148 }]);
    expect(data.daily_nutrition).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });

  it("should return a single sheet when sheet parameter is provided", () => {
    doGet(authedEvent({ sheet: "steps" }));
    const content = mockContentService.setContent.mock.calls[0][0];
    expect(JSON.parse(content)).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });

  it("should return an error when the sheet does not exist", () => {
    mockSpreadsheet.getSheetByName.mockReturnValue(null);
    doGet(authedEvent({ sheet: "unknown" }));
    const content = mockContentService.setContent.mock.calls[0][0];
    expect(JSON.parse(content)).toEqual({
      error: "sheet_not_found",
      sheet: "unknown",
    });
  });

  it("should return sheet_not_found when a known sheet is missing from the spreadsheet", () => {
    mockSpreadsheet.getSheetByName.mockReturnValue(null);
    doGet(authedEvent({ sheet: "steps" }));
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
    doGet(authedEvent({ sheet: "steps" }));
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
    doGet(authedEvent({ sheet: "steps", callback: "myCallback" }));
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JAVASCRIPT");
    expect(mockContentService.setContent).toHaveBeenCalledWith(
      "myCallback&&myCallback(" +
        JSON.stringify([{ date: "2026-08-23", steps: 6148 }]) +
        ");"
    );
  });

  it("should fall back to JSON for a callback that is not a plain identifier", () => {
    doGet(authedEvent({ sheet: "steps", callback: "alert(1)//" }));
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("JSON");
    expect(mockContentService.setContent).toHaveBeenCalledWith(
      JSON.stringify([{ date: "2026-08-23", steps: 6148 }])
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

  it("does not skip authentication when DEBUG is true", () => {
    scriptProperties.DEBUG = "true";
    doGet({ parameter: { sheet: "steps" } });
    expect(lastJson()).toEqual({ status: 401, error: "missing id token" });
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
    scriptProperties.SPREADSHEET_ID = "staging-file-id";
    SpreadsheetApp.openById.mockClear();
    doGet(authedEvent({ sheet: "steps" }));
    expect(SpreadsheetApp.openById).toHaveBeenCalledWith("staging-file-id");
    delete scriptProperties.SPREADSHEET_ID;
  });

  it("falls back to the built-in spreadsheet id when the property is unset", () => {
    SpreadsheetApp.openById.mockClear();
    doGet(authedEvent({ sheet: "steps" }));
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
  });

  it("reads sheets once and serves the second full request from cache", () => {
    doGet(authedEvent());
    const callsAfterFirst = mockRange.getValues.mock.calls.length;
    expect(callsAfterFirst).toBeGreaterThan(0);
    doGet(authedEvent());
    expect(mockRange.getValues.mock.calls.length).toBe(callsAfterFirst);
  });

  it("serves a single-sheet request from the cache warmed by a full request", () => {
    doGet(authedEvent());
    const calls = mockRange.getValues.mock.calls.length;
    doGet(authedEvent({ sheet: "steps" }));
    expect(mockRange.getValues.mock.calls.length).toBe(calls);
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    expect(JSON.parse(content)).toEqual([{ date: "2026-08-23", steps: 6148 }]);
  });
});

describe("period_summaries computed from daily_nutrition", () => {
  const dailyValues = [
    ["date", "intake_kcal", "balance_kcal", "completeness"],
    ["2026-11-02", 2000, -500, "完了"],
    ["2026-11-03", 2100, -600, "完了"],
    ["2026-11-04", 1900, -400, "完了"],
    ["2026-11-05", 2000, -500, "完了"],
    ["2026-11-06", 2200, -700, "完了"],
    ["2026-11-07", 1800, -300, "完了"],
    ["2026-11-08", 2000, -500, "完了"],
    ["2026-11-09T00:00:00+09:00", 2200, -900, "完了"],
    ["2026-11-10", 1955, -511, "途中（当日分）"],
  ];
  const dailySheet = {
    getDataRange: vi.fn().mockImplementation(() => ({
      getValues: vi.fn().mockImplementation(() => dailyValues.map((row) => [...row])),
    })),
  };

  beforeEach(() => {
    vi.clearAllMocks();
    cacheStore.clear();
    mockSpreadsheet.getSheetByName.mockImplementation((name) =>
      name === "daily_nutrition" ? dailySheet : mockSheet
    );
  });

  it("computes weekly summaries for a single-sheet request", () => {
    doGet(authedEvent({ sheet: "period_summaries" }));
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    const data = JSON.parse(content);
    expect(data).toHaveLength(2);
    expect(data[0]).toMatchObject({
      start_date: "2026-11-02",
      end_date: "2026-11-08",
      days: 7,
      average_intake_kcal: 2000,
      total_balance_kcal: -3500,
      average_balance_kcal: -500,
      completeness: "完了",
    });
    expect(data[1]).toMatchObject({
      start_date: "2026-11-09",
      end_date: "2026-11-09",
      days: 1,
      average_intake_kcal: 2200,
      total_balance_kcal: -900,
      completeness: "途中集計",
    });
  });

  it("includes the computed summaries in the full payload", () => {
    doGet(authedEvent());
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    const data = JSON.parse(content);
    expect(data.period_summaries).toHaveLength(2);
    expect(data.period_summaries[0].start_date).toBe("2026-11-02");
  });
});

describe("weekly_rss slot", () => {
  it("uses this Monday once 06:20 has passed", () => {
    const slot = weeklyRssSlot_(Date.parse("2026-11-16T07:00:00+09:00"));
    expect(slot.mondayKey).toBe("2026-11-16");
    expect(slot.weekStartKey).toBe("2026-11-09");
    expect(slot.weekEndKey).toBe("2026-11-15");
    expect(slot.atMs).toBe(Date.parse("2026-11-16T06:20:00+09:00"));
  });

  it("falls back to the previous Monday before 06:20 and on Sunday", () => {
    const early = weeklyRssSlot_(Date.parse("2026-11-16T05:00:00+09:00"));
    expect(early.mondayKey).toBe("2026-11-09");
    expect(early.weekStartKey).toBe("2026-11-02");
    expect(early.weekEndKey).toBe("2026-11-08");
    const sunday = weeklyRssSlot_(Date.parse("2026-11-15T12:00:00+09:00"));
    expect(sunday.mondayKey).toBe("2026-11-09");
  });

  it("resolves midweek to the current week's Monday", () => {
    const slot = weeklyRssSlot_(Date.parse("2026-11-18T12:00:00+09:00"));
    expect(slot.mondayKey).toBe("2026-11-16");
    expect(slot.weekStartKey).toBe("2026-11-09");
  });
});

describe("buildWeeklyRss_", () => {
  it("renders items with escaping and skips invalid ones", () => {
    const xml = buildWeeklyRss_([
      {
        title: "a<b",
        guid: "g1",
        pubDate: "Mon, 16 Nov 2026 21:20:00 GMT",
        descriptionText: "x & y",
      },
      { title: "no guid" },
      null,
    ]);
    expect(xml).toContain("<title>a&lt;b</title>");
    expect(xml).toContain("<description>x &amp; y</description>");
    expect(xml).not.toContain("no guid");
    expect(buildWeeklyRss_(null)).toContain("<channel>");
  });
});

describe("buildWeeklyReviewItem_", () => {
  const nowMs = Date.parse("2026-11-16T07:00:00+09:00");
  const data = {
    body_composition: [
      { date: "2018-09", weight_kg: 90.3 },
      { date: "2026-08-29", weight_kg: 95.7 },
      { date: "2026-10-17", weight_kg: 91.5, body_fat_percent: 28.6, fat_mass_kg: 26.2, muscle_mass_kg: 61.9, visceral_fat_level: 17, bmi: 32.4 },
      { date: "2026-11-14", weight_kg: 90.9, body_fat_percent: 28.0, fat_mass_kg: 25.5, muscle_mass_kg: 62.0, visceral_fat_level: 17, bmi: 32.2 },
    ],
    daily_nutrition: [
      { date: "2026-08-30", intake_kcal: 2000, expenditure_kcal: 3000, balance_kcal: -1000, completeness: "完了" },
      { date: "2026-11-09", intake_kcal: 1900, expenditure_kcal: 2800, balance_kcal: -900, completeness: "完了" },
      { date: "2026-11-10", intake_kcal: 2500, expenditure_kcal: 3100, balance_kcal: -600, protein_g: 80, fat_g: 90, carbohydrates_g: 240, salt_g: 12.5, completeness: "完了" },
      { date: "2026-11-11", intake_kcal: 2000, expenditure_kcal: 2900, balance_kcal: -900, completeness: "完了" },
      { date: "2026-11-12", intake_kcal: 2350, expenditure_kcal: 3250, balance_kcal: -900, completeness: "完了" },
      { date: "2026-11-13", intake_kcal: 2100, expenditure_kcal: 3100, balance_kcal: -1000, completeness: "完了" },
      { date: "2026-11-14", intake_kcal: 2250, expenditure_kcal: 3200, balance_kcal: -950, completeness: "完了" },
      { date: "2026-11-15", intake_kcal: 1955, expenditure_kcal: 2466, balance_kcal: -511, completeness: "途中（当日分）" },
    ],
    steps: [
      { date: "2026-11-09", steps: 9606, distance_walking_running_m: 6356.4, flights_climbed: 23, completeness: "完了" },
      { date: "2026-11-10", steps: 15407, distance_walking_running_m: 11643.3, flights_climbed: 19, completeness: "完了" },
      { date: "2026-11-15", steps: 8978, distance_walking_running_m: 5740.7, flights_climbed: 8, completeness: "途中（当日分）" },
    ],
    chat_activity: [
      { date: "2026-11-09", user_message_count: 120, engaged_minutes_est: 600, late_night_message_count: 5 },
      { date: "2026-11-10", user_message_count: 200, engaged_minutes_est: 900, late_night_message_count: 12 },
    ],
    sleep_estimates: [
      { date: "2026-11-10", sleep_start_est: "00:05", sleep_end_est: "03:05", duration_minutes_est: 180 },
      { date: "2026-11-11", sleep_start_est: "22:07", sleep_end_est: "04:20", duration_minutes_est: 373 },
    ],
    lab_results: [
      { date: "2026-11-12", item: "尿酸", value: 6.9, unit: "mg/dl", flag: "" },
      { date: "2026-11-12", item: "クレアチニン", value: 1.13, unit: "mg/dl", flag: "H" },
      { date: "2026-10-01", item: "尿酸", value: 6.0, unit: "mg/dl", flag: "" },
    ],
    goals_settings: [
      { item: "開始時体重", value: "99.7", unit_or_date: "kg / 2026-07-25" },
      { item: "中間目標", value: "90" },
      { item: "最終目標", value: "80" },
      { item: "エネルギー目標", value: "2000" },
      { item: "タンパク質目標", value: "120" },
      { item: "塩分目標", value: "6" },
    ],
  };

  it("builds the full weekly review for 2026-11-09..15", () => {
    const item = buildWeeklyReviewItem_(data, nowMs);
    expect(item.title).toBe("11/9〜11/15 健康ウィークリー");
    expect(item.guid).toBe("health-weekly-2026-11-16");
    expect(item.pubDate).toBe(new Date(Date.parse("2026-11-16T06:20:00+09:00")).toUTCString());
    const text = item.descriptionText;
    expect(text).toContain("11/9〜11/15の週次レビュー。");
    expect(text).toContain("■ 体重・体組成");
    expect(text).toContain("11/14（土） 体重 90.9 kg・体脂肪率 28.0%・脂肪量 25.5 kg・筋肉量 62.0 kg・内臓脂肪レベル 17・BMI 32.2");
    expect(text).toContain("最新の実測：11/14 90.9 kg（前回 10/17 比 -0.6 kg）");
    expect(text).toContain("開始時（7/25 99.7 kg）から累計 -8.8 kg（-8.8%）");
    expect(text).toContain("中間目標 90.0 kg まで あと 0.9 kg／最終目標 80.0 kg まで あと 10.9 kg");
    expect(text).toContain("実測換算：体重 1 kg あたり 8,750 kcal（10/17〜11/14 の累積収支 -5,250 kcal から算出）");
    expect(text).toContain("直近30日（完了日）の平均収支：-893 kcal/日");
    expect(text).toContain("このペースでの30日後予測：87.8 kg");
    expect(text).toContain("中間目標 90.0 kg の到達予測：2026/11/23");
    expect(text).toContain("最終目標 80.0 kg の到達予測：2027/03/01");
    expect(text).toContain("■ 栄養（からサポ）");
    expect(text).toContain("11/9（月） 摂取 1,900／消費 2,800／収支 -900 kcal");
    expect(text).toContain("栄養素：タンパク質 80 g・脂質 90 g・炭水化物 240 g・塩分 12.5 g");
    expect(text).toContain("11/15（日） 摂取 1,955／消費 2,466／収支 -511 kcal（途中（当日分））");
    expect(text).toContain("週の合計（完了 6 日）：摂取 13,100／消費 18,350／収支 -5,250 kcal");
    expect(text).toContain("栄養指導の目標：エネルギー 2,000 kcal/日・タンパク質 120 g/日・塩分 6.0 g/日");
    expect(text).toContain("■ 歩数");
    expect(text).toContain("11/9（月） 9,606 歩（6,356 m・23 階）");
    expect(text).toContain("11/11（水） 記録なし");
    expect(text).toContain("週の合計：33,991 歩");
    expect(text).toContain("■ 活動時間");
    expect(text).toContain("11/9（月） 10時間00分（活動 120 回・深夜帯 5 回）");
    expect(text).toContain("■ 睡眠（推定）");
    expect(text).toContain("11/10（火） 00:05〜03:05（3時間00分）");
    expect(text).toContain("■ 検査結果");
    expect(text).toContain("・クレアチニン 1.13 mg/dl（H）");
    expect(text).not.toContain("2026-10-01");
  });

  it("tolerates empty data", () => {
    const item = buildWeeklyReviewItem_({}, nowMs);
    expect(item.descriptionText).toContain("この週の実測はありません。");
    expect(item.descriptionText).toContain("実測換算はデータ不足のため");
    expect(item.descriptionText).toContain("この週の新しい検査結果はありません。");
  });

  it("adds the predicted weight at output time from balances after the latest measurement", () => {
    const item = buildWeeklyReviewItem_(
      {
        body_composition: [{ date: "2026-11-07", weight_kg: 92.0 }],
        daily_nutrition: [
          { date: "2026-11-08", intake_kcal: 2000, expenditure_kcal: 2500, balance_kcal: -500, completeness: "完了" },
          { date: "2026-11-09", intake_kcal: 2000, expenditure_kcal: 2500, balance_kcal: -500, completeness: "完了" },
          { date: "2026-11-10", intake_kcal: 2000, expenditure_kcal: 2500, balance_kcal: -500, completeness: "途中（当日分）" },
        ],
        steps: [],
        chat_activity: [],
        sleep_estimates: [],
        lab_results: [],
        goals_settings: [],
      },
      nowMs,
    );
    // 92.0 kg on 11/7 plus -1,000 kcal of completed balances at the
    // 7,700 fallback factor -> 91.9 kg. The 途中 row is excluded.
    expect(item.descriptionText).toContain(
      "出力時点の推定体重：91.9 kg（11/9 時点。最新実測 11/7 の 92.0 kg から、完了した収支の累計 -1,000 kcal を実測換算で換算）",
    );
    expect(item.descriptionText).toContain("推定体重は 91.9 kg。");
  });
});

describe("doGet weekly_rss", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    cacheStore.clear();
    delete scriptProperties.WEEKLY_RSS_TOKEN;
    const emptySheet = {
      getDataRange: vi.fn().mockImplementation(() => ({
        getValues: vi.fn().mockReturnValue([["date"]]),
      })),
    };
    mockSpreadsheet.getSheetByName.mockImplementation(() => emptySheet);
  });

  it("refuses when the token property is unset", () => {
    doGet({ parameter: { action: "weekly_rss", token: "anything" } });
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    expect(JSON.parse(content)).toMatchObject({ ok: false, status: 403 });
  });

  it("refuses a mismatched token", () => {
    scriptProperties.WEEKLY_RSS_TOKEN = "expected-token";
    doGet({ parameter: { action: "weekly_rss", token: "wrong" } });
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    expect(JSON.parse(content)).toMatchObject({ ok: false, status: 403 });
  });

  it("serves the RSS feed for the right token", () => {
    scriptProperties.WEEKLY_RSS_TOKEN = "expected-token";
    doGet({ parameter: { action: "weekly_rss", token: "expected-token" } });
    expect(mockContentService.setMimeType).toHaveBeenCalledWith("XML");
    const content = mockContentService.setContent.mock.calls.at(-1)[0];
    expect(content).toContain('<rss version="2.0">');
    expect(content).toContain("健康ウィークリー");
    expect(content).toContain("<item>");
  });
});
