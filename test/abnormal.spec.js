import { describe, expect, it, vi } from "vitest";
import { buildWeeklyReviewItem_, buildWeeklyRss_, doGet } from "../src/Code.js";

// Abnormal-case tests: the "input from outside" for health-gas is the
// health spreadsheet itself. getValues() yields scalar cells (string /
// number / boolean / Date / null), but imports and manual edits can put
// text in numeric columns, booleans in measurement cells, non-existent
// dates in date columns, duplicated dates, and negative counts. One
// deformed cell must not invent measurements, leak NaN/Infinity into
// the payload or the weekly RSS, or kill the whole response.
//
// All figures below are synthetic.

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
  AVAILABLE_GMAILS: "owner@example.com",
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
        email: "owner@example.com",
        sub: "1",
      }),
    ),
  })),
};

const DAILY_HEADERS = ["date", "intake_kcal", "expenditure_kcal", "balance_kcal", "completeness"];
const periodSummaries = (rows) => {
  cacheStore.clear();
  sheetValues = [DAILY_HEADERS, ...rows];
  doGet({ parameter: { sheet: "period_summaries", id_token: "good" } });
  return JSON.parse(global.ContentService.setContent.mock.calls.at(-1)[0]);
};
const sheetRows = (headers, rows) => {
  cacheStore.clear();
  sheetValues = [headers, ...rows];
  doGet({ parameter: { sheet: "steps", id_token: "good" } });
  return JSON.parse(global.ContentService.setContent.mock.calls.at(-1)[0]);
};

describe("period_summaries with deformed daily_nutrition rows", () => {
  it("drops a completed row whose expenditure is not numeric instead of leaking NaN", () => {
    // On the original code average_expenditure_kcal came out as NaN
    // (null after JSON serialization).
    expect(periodSummaries([["2026-11-02", 2000, "abc", -500, "完了"]])).toEqual([]);
  });

  it("drops a completed row whose kcal fields are booleans instead of counting them as 1", () => {
    expect(periodSummaries([["2026-11-02", true, 2500, true, "完了"]])).toEqual([]);
  });

  it("drops a completed row whose kcal field is a single-element array", () => {
    expect(periodSummaries([["2026-11-02", [2000], 2500, -500, "完了"]])).toEqual([]);
  });

  it("drops a completed row whose kcal field is a blank string", () => {
    expect(periodSummaries([["2026-11-02", " ", 2500, -500, "完了"]])).toEqual([]);
  });

  it("still counts numeric-string kcal values (sheet cells are often text)", () => {
    const out = periodSummaries([["2026-11-02", "2000", "2500", "-500", "完了"]]);
    expect(out).toHaveLength(1);
    expect(out[0]).toMatchObject({
      average_intake_kcal: 2000,
      average_expenditure_kcal: 2500,
      total_balance_kcal: -500,
    });
  });

  it("drops a completed row dated on a day that does not exist", () => {
    expect(periodSummaries([["2026-02-30", 2000, 2500, -500, "完了"]])).toEqual([]);
    expect(periodSummaries([["2023-02-29", 2000, 2500, -500, "完了"]])).toEqual([]);
  });

  it("counts a duplicated date once, keeping the last row", () => {
    const out = periodSummaries([
      ["2026-11-02", 2000, 2500, -500, "完了"],
      ["2026-11-02", 1000, 1500, -250, "完了"],
    ]);
    expect(out).toHaveLength(1);
    expect(out[0]).toMatchObject({
      days: 1,
      average_intake_kcal: 1000,
      total_balance_kcal: -250,
    });
  });

  it("survives a null row in the sheet values", () => {
    const out = periodSummaries([null, ["2026-11-02", 2000, 2500, -500, "完了"]]);
    expect(out).toHaveLength(1);
    expect(out[0].days).toBe(1);
  });
});

describe("sheet shaping with deformed rows", () => {
  it("skips non-array rows instead of throwing", () => {
    const out = sheetRows(["date", "steps"], [null, 42, ["2026-11-09", 100]]);
    expect(out).toEqual([{ date: "2026-11-09", steps: 100 }]);
  });

  it("reads cells past the end of a short row as null", () => {
    const out = sheetRows(["date", "steps"], [["2026-11-09"]]);
    expect(out).toEqual([{ date: "2026-11-09", steps: null }]);
  });
});

// --- Weekly RSS builder ----------------------------------------------
const nowMs = Date.parse("2026-11-16T07:00:00+09:00"); // slot week 2026-11-09..15
const emptyData = {
  body_composition: [],
  daily_nutrition: [],
  steps: [],
  chat_activity: [],
  sleep_estimates: [],
  lab_results: [],
  goals_settings: [],
};
const build = (patch, at = nowMs) => buildWeeklyReviewItem_({ ...emptyData, ...patch }, at);

describe("buildWeeklyReviewItem_ with deformed measurements", () => {
  it("treats a boolean weight as no measurement, not 1.0 kg", () => {
    const item = build({ body_composition: [{ date: "2026-11-14", weight_kg: true }] });
    expect(item.descriptionText).toContain("この週の実測はありません。");
    expect(item.descriptionText).not.toContain("1.0 kg");
  });

  it("treats an array weight as no measurement", () => {
    const item = build({ body_composition: [{ date: "2026-11-14", weight_kg: [90.5] }] });
    expect(item.descriptionText).toContain("この週の実測はありません。");
  });

  it("treats a Date weight as no measurement, not epoch milliseconds", () => {
    const item = build({
      body_composition: [{ date: "2026-11-14", weight_kg: new Date(Date.UTC(2026, 0, 1)) }],
    });
    expect(item.descriptionText).toContain("この週の実測はありません。");
    expect(item.descriptionText).not.toContain("1767225600000");
  });

  it("still accepts a numeric-string weight", () => {
    const item = build({ body_composition: [{ date: "2026-11-14", weight_kg: "90.5" }] });
    expect(item.descriptionText).toContain("体重 90.5 kg");
  });

  it("treats a boolean step count as no record, not 1 step", () => {
    const item = build({ steps: [{ date: "2026-11-09", steps: true }] });
    expect(item.descriptionText).toContain("11/9（月） 記録なし");
    expect(item.descriptionText).not.toContain("合計 1 歩");
  });

  it("treats negative steps as no record instead of printing a negative total", () => {
    const item = build({ steps: [{ date: "2026-11-09", steps: -500 }] });
    expect(item.descriptionText).toContain("11/9（月） 記録なし");
    expect(item.descriptionText).not.toContain("-500");
  });

  it("treats a negative sleep duration as no record instead of a negative duration text", () => {
    const item = build({
      sleep_estimates: [
        { date: "2026-11-10", sleep_start_est: "23:00", sleep_end_est: "04:00", duration_minutes_est: -30 },
      ],
    });
    expect(item.descriptionText).toContain("11/10（火） 記録なし");
    expect(item.descriptionText).not.toContain("-1時間-30分");
  });

  it("treats a negative activity duration as no record", () => {
    const item = build({
      chat_activity: [{ date: "2026-11-09", engaged_minutes_est: -90, user_message_count: 3 }],
    });
    expect(item.descriptionText).toContain("11/9（月） 記録なし");
    expect(item.descriptionText).not.toContain("-2時間-30分");
  });

  it("ignores a body row dated on a day that does not exist", () => {
    const febNow = Date.parse("2026-03-02T07:00:00+09:00"); // slot week 2026-02-23..03-01
    const item = build({ body_composition: [{ date: "2026-02-30", weight_kg: 90 }] }, febNow);
    expect(item.descriptionText).toContain("この週の実測はありません。");
    expect(item.descriptionText).not.toContain("2/30");
  });

  it("ignores a lab row dated on a day that does not exist", () => {
    const febNow = Date.parse("2026-03-02T07:00:00+09:00");
    const item = build(
      { lab_results: [{ date: "2026-02-30", item: "X", value: 1, unit: "mg", flag: "" }] },
      febNow,
    );
    expect(item.descriptionText).toContain("この週の新しい検査結果はありません。");
    expect(item.descriptionText).not.toContain("2/30");
  });

  it("falls back to the theoretical factor when the empirical factor rounds to 0 kcal/kg", () => {
    // -1 kcal of cumulative balance over a 3 kg loss rounds to a
    // 0 kcal/kg factor; dividing the forecast by it printed
    // "-Infinity kg" on the original code.
    const item = build({
      body_composition: [
        { date: "2026-11-03", weight_kg: 93 },
        { date: "2026-11-14", weight_kg: 90 },
      ],
      daily_nutrition: [
        { date: "2026-11-03", intake_kcal: 2000, expenditure_kcal: 2001, balance_kcal: -1, completeness: "完了" },
        { date: "2026-11-10", intake_kcal: 2000, expenditure_kcal: 2001, balance_kcal: -1, completeness: "完了" },
      ],
    });
    expect(item.descriptionText).toContain("実測換算はデータ不足のため");
    expect(item.descriptionText).not.toContain("Infinity");
    expect(item.descriptionText).not.toContain("0 kcal（");
  });

  it("does not divide by a zero start weight when printing the cumulative change", () => {
    // Found by the property suite over well-formed data: a 開始時体重
    // of 0 made the percentage Infinity.
    const item = build({
      body_composition: [{ date: "2026-11-14", weight_kg: 90 }],
      goals_settings: [{ item: "開始時体重", value: 0, unit_or_date: "kg / 2026-07-25" }],
    });
    expect(item.descriptionText).not.toContain("Infinity");
    expect(item.descriptionText).not.toContain("開始時（");
  });

  it("does not let the one-decimal formatter overflow a finite percentage into Infinity", () => {
    // Regression (post-merge CI failure of PR #31, seed-dependent):
    // a denormal-tiny start weight makes the cumulative percentage
    // exactly Number.MAX_VALUE — finite, so the isFinite guard on the
    // percentage passed — but weeklyRssFmt1_ multiplied it by 10 for
    // rounding, overflowed, and printed "+Infinity%".
    const item = build({
      body_composition: [{ date: "2020-01-01", weight_kg: 30 }],
      goals_settings: [{ item: "開始時体重", value: 1.6688053938804013e-305 }],
    });
    expect(item.descriptionText).not.toContain("Infinity");
    expect(item.descriptionText).toContain("開始時（");
  });

  it("omits a goal forecast whose date overflows instead of printing a NaN date", () => {
    // Found by the property suite: a near-zero average balance makes
    // the day count overflow the Date range (NaN/NaN/NaN on the
    // original code).
    const item = build({
      body_composition: [{ date: "2026-11-14", weight_kg: 90 }],
      daily_nutrition: [
        { date: "2026-11-10", intake_kcal: 2000, expenditure_kcal: 2000, balance_kcal: -0.001, completeness: "完了" },
      ],
      goals_settings: [{ item: "最終目標", value: 80 }],
    });
    expect(item.descriptionText).not.toContain("NaN");
    expect(item.descriptionText).not.toContain("到達予測");
  });

  it("excludes completed rows with an unparseable balance from the 30-day average", () => {
    const rows = [];
    for (let i = 0; i < 10; i += 1) {
      rows.push({
        date: `2026-11-${String(5 + i).padStart(2, "0")}`,
        intake_kcal: 2000,
        expenditure_kcal: 3000,
        balance_kcal: -1000,
        completeness: "完了",
      });
    }
    rows.push({
      date: "2026-11-15",
      intake_kcal: 2000,
      expenditure_kcal: 3000,
      balance_kcal: "abc",
      completeness: "完了",
    });
    const item = build({
      daily_nutrition: rows,
      body_composition: [{ date: "2026-11-14", weight_kg: 90 }],
    });
    // The broken row counted as 0 on the original code (-909 average).
    expect(item.descriptionText).toContain("直近30日（完了日）の平均収支：-1,000 kcal/日");
  });

  it("tolerates a non-array sheet value and null data", () => {
    expect(() => build({ body_composition: "x" })).not.toThrow();
    expect(build({ body_composition: "x" }).descriptionText).toContain("この週の実測はありません。");
    expect(() => buildWeeklyReviewItem_(null, nowMs)).not.toThrow();
    expect(buildWeeklyReviewItem_(null, nowMs).descriptionText).toContain(
      "この週の実測はありません。",
    );
  });

  it("tolerates unstringifiable text fields without throwing or printing [object Object]", () => {
    const item = build({
      sleep_estimates: [
        {
          date: "2026-11-10",
          sleep_start_est: { toString: [] },
          sleep_end_est: "04:00",
          duration_minutes_est: 300,
        },
      ],
      lab_results: [
        { date: "2026-11-12", item: "X", value: { toString: [] }, unit: "mg", flag: "" },
      ],
    });
    expect(item.descriptionText).not.toContain("[object Object]");
    expect(item.descriptionText).toContain("11/10（火）");
  });

  it("never prints NaN, Infinity or undefined for a mixed deformed dataset", () => {
    const item = build({
      body_composition: [
        { date: "2026-11-14", weight_kg: "abc", body_fat_percent: true },
        { date: "2026-02-30", weight_kg: 90 },
      ],
      daily_nutrition: [
        { date: "2026-11-09", intake_kcal: [], expenditure_kcal: {}, balance_kcal: "x", completeness: "完了" },
      ],
      steps: [{ date: "2026-11-09", steps: -1 }],
      sleep_estimates: [{ date: "2026-11-10", duration_minutes_est: "x" }],
      lab_results: [{ date: "2026-11-12", item: "X", value: "1.5", unit: "", flag: "" }],
      goals_settings: [{ item: "最終目標", value: true }],
    });
    for (const bad of ["NaN", "Infinity", "undefined", "[object Object]"]) {
      expect(item.descriptionText).not.toContain(bad);
      expect(item.title).not.toContain(bad);
    }
  });
});

describe("buildWeeklyRss_ escaping", () => {
  const item = (descriptionText) => ({
    title: "t",
    guid: "g",
    pubDate: "Mon, 16 Nov 2026 21:20:00 GMT",
    descriptionText,
  });

  it("replaces lone surrogates and C0 controls so the feed stays well-formed XML", () => {
    const xml = buildWeeklyRss_([item("a\uD800b\x07c")]);
    expect(xml).not.toContain("\uD800");
    expect(xml).not.toContain("\x07");
    expect(xml).toContain("a\uFFFDbc");
  });

  it("keeps paired emoji and newlines intact", () => {
    const xml = buildWeeklyRss_([item("x🎉y\nz")]);
    expect(xml).toContain("x🎉y\nz");
  });
});
