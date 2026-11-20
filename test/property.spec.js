import { describe, expect, it, vi } from "vitest";
import * as fc from "fast-check";
import { buildWeeklyReviewItem_, buildWeeklyRss_, doGet, weeklyRssSlot_ } from "../src/Code.js";

// Property-based tests (fast-check) over the same outside-input
// boundary as test/abnormal.spec.js: spreadsheet cell values and the
// rows built from them.
//
// Generator domains, and why:
// - Cell values are drawn from the JSON domain (strings, finite
//   numbers, booleans, null, arrays, plain objects) plus one explicit
//   unstringifiable object ({ toString: [] }), which JSON can produce
//   and which must never make a reader throw. Sheet cells themselves
//   are scalars, so this domain is a superset of the real one.
// - Strings are built from whole code points (fc.string with
//   fc.char16bits would invent lone surrogates that real sheet text
//   cannot contain); lone-surrogate handling is pinned separately in
//   the abnormal suite for the XML boundary, where it matters.
// - Dates are either real calendar dates (fc.date with
//   noInvalidDate, formatted in UTC) or a fixed set of misshapen keys.
// - All figures are synthetic.

// --- GAS mocks (mirroring test/Code.spec.js) -------------------------
const cacheStore = new Map();
let sheetValues = [];
const mockSheet = {
  getDataRange: vi.fn().mockImplementation(() => ({
    getValues: vi
      .fn()
      .mockImplementation(() => sheetValues.map((r) => (Array.isArray(r) ? [...r] : r))),
  })),
};
const mockSpreadsheet = { getSheetByName: vi.fn().mockImplementation(() => mockSheet) };
global.SpreadsheetApp = { openById: vi.fn().mockReturnValue(mockSpreadsheet) };
global.ContentService = {
  createTextOutput: vi.fn().mockReturnThis(),
  setMimeType: vi.fn().mockReturnThis(),
  setContent: vi.fn(),
  MimeType: { JSON: "JSON", JAVASCRIPT: "JAVASCRIPT", XML: "XML" },
};
global.Utilities = {
  formatDate: vi.fn().mockReturnValue("2026-08-23T00:00:00+09:00"),
  DigestAlgorithm: { SHA_256: "SHA_256" },
  computeDigest: vi.fn().mockImplementation(() => Array.from({ length: 32 }, () => 0)),
  getUuid: vi.fn().mockReturnValue("00000000-0000-4000-8000-000000000000"),
};
global.CacheService = {
  getScriptCache: vi.fn().mockImplementation(() => ({
    get: vi.fn().mockImplementation((k) => (cacheStore.has(k) ? cacheStore.get(k) : null)),
    getAll: vi.fn().mockImplementation((keys) => {
      const out = {};
      keys.forEach((k) => {
        if (cacheStore.has(k)) out[k] = cacheStore.get(k);
      });
      return out;
    }),
    putAll: vi.fn().mockImplementation((entries) => {
      Object.entries(entries).forEach(([k, v]) => cacheStore.set(k, v));
    }),
  })),
};
global.Logger = { log: vi.fn() };
const scriptProperties = {
  AVAILABLE_GMAILS: "zegenvs@gmail.com",
  GOOGLE_OAUTH_CLIENT_ID: "test-client-id",
};
global.PropertiesService = {
  getScriptProperties: vi.fn().mockImplementation(() => ({
    getProperty: vi.fn().mockImplementation((key) => scriptProperties[key] ?? null),
  })),
};
global.UrlFetchApp = {
  fetch: vi.fn().mockImplementation(() => ({
    getResponseCode: vi.fn().mockReturnValue(200),
    getContentText: vi.fn().mockReturnValue(
      JSON.stringify({
        iss: "https://accounts.google.com",
        aud: "test-client-id",
        exp: String(Math.floor(Date.now() / 1000) + 3600),
        email_verified: "true",
        email: "zegenvs@gmail.com",
        sub: "1",
      }),
    ),
  })),
};

// --- Generators --------------------------------------------------------
const unstringifiable = { toString: [] };
const jsonValue = fc.oneof(
  fc.string({ maxLength: 12 }),
  fc.integer({ min: -100000, max: 100000 }),
  fc.double({ min: -1e6, max: 1e6, noNaN: true }),
  fc.boolean(),
  fc.constant(null),
  fc.array(fc.integer({ min: -100, max: 100 }), { maxLength: 3 }),
  fc.constant({}),
  fc.constant(unstringifiable),
);
const realDateKey = fc
  .date({ min: new Date(Date.UTC(2020, 0, 1)), max: new Date(Date.UTC(2030, 11, 31)), noInvalidDate: true })
  .map((d) => d.toISOString().slice(0, 10));
const misshapenDateKey = fc.constantFrom(
  "2026-02-30",
  "2023-02-29",
  "2026-13-01",
  "2026-1-1",
  "not-a-date",
  "",
  "2026-11-09T00:00:00+09:00",
);
const anyDateKey = fc.oneof(realDateKey, misshapenDateKey);
const isRealDateKey = (key) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) return false;
  const [y, m, d] = key.split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
};
const hasLoneSurrogate = (s) => {
  for (let i = 0; i < s.length; i += 1) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = s.charCodeAt(i + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return true;
      i += 1;
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      return true;
    }
  }
  return false;
};
const BAD_TOKENS = ["NaN", "Infinity", "undefined", "[object Object]"];

const nowMs = Date.parse("2026-11-16T07:00:00+09:00");

describe("period_summaries properties", () => {
  const DAILY_HEADERS = ["date", "intake_kcal", "expenditure_kcal", "balance_kcal", "completeness"];
  const runSummaries = (rows) => {
    cacheStore.clear();
    sheetValues = [DAILY_HEADERS, ...rows];
    doGet({ parameter: { sheet: "period_summaries", id_token: "good" } });
    return JSON.parse(global.ContentService.setContent.mock.calls.at(-1)[0]);
  };
  const dailyRow = fc.record({
    date: anyDateKey,
    intake: jsonValue,
    expenditure: jsonValue,
    balance: jsonValue,
    completeness: fc.constantFrom("完了", "途中（当日分）", "", "x"),
  });

  it("never throws and every reported figure is finite with real dates", () => {
    fc.assert(
      fc.property(fc.array(dailyRow, { maxLength: 20 }), (rows) => {
        const values = rows.map((r) => [r.date, r.intake, r.expenditure, r.balance, r.completeness]);
        const out = runSummaries(values);
        expect(Array.isArray(out)).toBe(true);
        for (const s of out) {
          for (const k of [
            "average_intake_kcal",
            "average_expenditure_kcal",
            "total_balance_kcal",
            "average_balance_kcal",
          ]) {
            expect(Number.isFinite(s[k])).toBe(true);
          }
          expect(isRealDateKey(s.start_date)).toBe(true);
          expect(isRealDateKey(s.end_date)).toBe(true);
          expect(s.start_date <= s.end_date).toBe(true);
          expect(s.days).toBeGreaterThanOrEqual(1);
          expect(s.days).toBeLessThanOrEqual(7);
        }
      }),
      { numRuns: 100 },
    );
  });

  it("conserves totals over well-formed completed rows", () => {
    const wellFormed = fc.record({
      date: realDateKey,
      intake: fc.integer({ min: 0, max: 5000 }),
      expenditure: fc.integer({ min: 0, max: 5000 }),
      balance: fc.integer({ min: -3000, max: 3000 }),
    });
    fc.assert(
      fc.property(fc.array(wellFormed, { minLength: 1, maxLength: 21 }), (rows) => {
        // Distinct dates: the last row for a date is the counted one.
        const byDate = new Map();
        rows.forEach((r) => byDate.set(r.date, r));
        const counted = [...byDate.values()];
        const values = rows.map((r) => [r.date, r.intake, r.expenditure, r.balance, "完了"]);
        const out = runSummaries(values);
        const totalDays = out.reduce((sum, s) => sum + s.days, 0);
        const totalBalance = out.reduce((sum, s) => sum + s.total_balance_kcal, 0);
        expect(totalDays).toBe(counted.length);
        expect(totalBalance).toBe(counted.reduce((sum, r) => sum + r.balance, 0));
      }),
      { numRuns: 100 },
    );
  });
});

describe("sheet shaping properties", () => {
  it("never throws on arbitrary sheet values and returns plain row objects", () => {
    const cell = fc.oneof(fc.string({ maxLength: 8 }), fc.integer(), fc.boolean(), fc.constant(null));
    const row = fc.oneof(
      fc.array(cell, { maxLength: 4 }),
      fc.constant(null),
      fc.integer(),
      fc.string({ maxLength: 4 }),
    );
    fc.assert(
      fc.property(fc.array(row, { maxLength: 8 }), (rows) => {
        cacheStore.clear();
        sheetValues = [["date", "steps"], ...rows];
        doGet({ parameter: { sheet: "steps", id_token: "good" } });
        const out = JSON.parse(global.ContentService.setContent.mock.calls.at(-1)[0]);
        expect(Array.isArray(out)).toBe(true);
        for (const entry of out) {
          expect(typeof entry).toBe("object");
          expect(entry).not.toBeNull();
        }
      }),
      { numRuns: 100 },
    );
  });
});

describe("weekly RSS builder properties", () => {
  const rowOf = (fields) =>
    fc.record(fields, { requiredKeys: [] });
  const dataArb = fc.record(
    {
      body_composition: fc.array(
        rowOf({ date: anyDateKey, weight_kg: jsonValue, body_fat_percent: jsonValue, bmi: jsonValue }),
        { maxLength: 6 },
      ),
      daily_nutrition: fc.array(
        rowOf({
          date: anyDateKey,
          intake_kcal: jsonValue,
          expenditure_kcal: jsonValue,
          balance_kcal: jsonValue,
          completeness: fc.constantFrom("完了", "途中（当日分）", ""),
          protein_g: jsonValue,
        }),
        { maxLength: 10 },
      ),
      steps: fc.array(rowOf({ date: anyDateKey, steps: jsonValue }), { maxLength: 6 }),
      chat_activity: fc.array(
        rowOf({ date: anyDateKey, engaged_minutes_est: jsonValue, user_message_count: jsonValue }),
        { maxLength: 6 },
      ),
      sleep_estimates: fc.array(
        rowOf({ date: anyDateKey, duration_minutes_est: jsonValue, sleep_start_est: jsonValue, sleep_end_est: jsonValue }),
        { maxLength: 6 },
      ),
      lab_results: fc.array(
        rowOf({ date: anyDateKey, item: jsonValue, value: jsonValue, unit: jsonValue, flag: jsonValue }),
        { maxLength: 6 },
      ),
      goals_settings: fc.array(rowOf({ item: jsonValue, value: jsonValue, unit_or_date: jsonValue }), {
        maxLength: 6,
      }),
    },
    { requiredKeys: [] },
  );

  it("never throws and never prints NaN, Infinity, undefined or [object Object]", () => {
    fc.assert(
      fc.property(dataArb, (data) => {
        const item = buildWeeklyReviewItem_(data, nowMs);
        expect(typeof item.title).toBe("string");
        expect(item.guid).toMatch(/^health-weekly-\d{4}-\d{2}-\d{2}$/);
        for (const bad of BAD_TOKENS) {
          expect(item.descriptionText).not.toContain(bad);
          expect(item.title).not.toContain(bad);
        }
        expect(hasLoneSurrogate(item.descriptionText)).toBe(false);
      }),
      { numRuns: 100 },
    );
  });

  it("keeps every printed figure finite over well-formed numeric data", () => {
    const num = fc.double({ min: -5000, max: 5000, noNaN: true });
    const wellFormed = fc.record({
      body_composition: fc.array(
        fc.record({ date: realDateKey, weight_kg: fc.double({ min: 30, max: 200, noNaN: true }) }),
        { maxLength: 5 },
      ),
      daily_nutrition: fc.array(
        fc.record({
          date: realDateKey,
          intake_kcal: num,
          expenditure_kcal: num,
          balance_kcal: num,
          completeness: fc.constantFrom("完了", "途中（当日分）"),
        }),
        { maxLength: 12 },
      ),
      steps: fc.array(fc.record({ date: realDateKey, steps: fc.integer({ min: 0, max: 30000 }) }), {
        maxLength: 5,
      }),
      chat_activity: fc.array(
        fc.record({ date: realDateKey, engaged_minutes_est: fc.integer({ min: 0, max: 1440 }) }),
        { maxLength: 5 },
      ),
      sleep_estimates: fc.array(
        fc.record({ date: realDateKey, duration_minutes_est: fc.integer({ min: 0, max: 1440 }) }),
        { maxLength: 5 },
      ),
      lab_results: fc.array(fc.record({ date: realDateKey, item: fc.string({ maxLength: 6 }), value: num }), {
        maxLength: 5,
      }),
      goals_settings: fc.array(
        fc.record({ item: fc.constantFrom("最終目標", "中間目標", "開始時体重"), value: num }),
        { maxLength: 3 },
      ),
    });
    fc.assert(
      fc.property(wellFormed, (data) => {
        const item = buildWeeklyReviewItem_(data, nowMs);
        for (const bad of BAD_TOKENS) {
          expect(item.descriptionText).not.toContain(bad);
        }
      }),
      { numRuns: 100 },
    );
  });
});

describe("buildWeeklyRss_ properties", () => {
  it("always emits well-formed XML text (no lone surrogates, no forbidden controls)", () => {
    const text = fc.string({ maxLength: 40 });
    fc.assert(
      fc.property(
        fc.array(
          fc.record({
            title: text,
            guid: fc.string({ minLength: 1, maxLength: 12 }),
            pubDate: fc.constant("Mon, 16 Nov 2026 21:20:00 GMT"),
            descriptionText: text,
          }),
          { maxLength: 4 },
        ),
        (items) => {
          const xml = buildWeeklyRss_(items);
          expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8"?>')).toBe(true);
          expect(hasLoneSurrogate(xml)).toBe(false);
          // No C0 controls other than tab / LF / CR anywhere.
          expect(/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(xml)).toBe(false);
        },
      ),
      { numRuns: 100 },
    );
  });

  it("escapes markup characters in item titles", () => {
    fc.assert(
      fc.property(fc.string({ maxLength: 30 }), (title) => {
        const xml = buildWeeklyRss_([
          { title, guid: "g", pubDate: "Mon, 16 Nov 2026 21:20:00 GMT", descriptionText: "" },
        ]);
        const itemTitle = xml.split("<item>")[1].split("</title>")[0].split("<title>")[1];
        expect(itemTitle).not.toContain("<");
        expect(itemTitle).not.toContain(">");
      }),
      { numRuns: 100 },
    );
  });
});

describe("weeklyRssSlot_ properties", () => {
  it("always resolves to a real Monday slot covering the previous week", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: Date.UTC(2020, 0, 1), max: Date.UTC(2030, 11, 31) }),
        (ms) => {
          const slot = weeklyRssSlot_(ms);
          expect(isRealDateKey(slot.mondayKey)).toBe(true);
          expect(new Date(Date.parse(slot.mondayKey + "T00:00:00Z")).getUTCDay()).toBe(1);
          expect(isRealDateKey(slot.weekStartKey)).toBe(true);
          expect(isRealDateKey(slot.weekEndKey)).toBe(true);
          expect(slot.weekStartKey < slot.weekEndKey).toBe(true);
          expect(slot.weekEndKey < slot.mondayKey).toBe(true);
          expect(slot.atMs).toBe(Date.parse(slot.mondayKey + "T06:20:00+09:00"));
          expect(slot.atMs <= ms).toBe(true);
        },
      ),
      { numRuns: 100 },
    );
  });
});
