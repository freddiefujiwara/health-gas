import { afterEach, describe, expect, it, vi } from "vitest";
import { parsePresenceTimestamp_, replaceHomeSegments_, replaceWorkSegments_, summarizeWorkDay_ } from "../src/Code.js";

const ev = (timestamp, action, place = "work") => ({ timestamp, person: "freddiefujiwara", place, action });

describe("parsePresenceTimestamp_", () => {
  it("parses ISO strings with an offset and epoch numbers", () => {
    expect(parsePresenceTimestamp_("2026-11-21T09:00:00+09:00")).toBe(Date.parse("2026-11-21T00:00:00Z"));
    expect(parsePresenceTimestamp_(1763686800000)).toBe(1763686800000);
  });

  it("reads naive strings as JST", () => {
    expect(parsePresenceTimestamp_("2026-11-21 09:00:00")).toBe(Date.parse("2026-11-21T00:00:00Z"));
    expect(parsePresenceTimestamp_("2026-11-21T09:00")).toBe(Date.parse("2026-11-21T00:00:00Z"));
  });

  it("returns NaN for unusable values", () => {
    expect(Number.isNaN(parsePresenceTimestamp_(null))).toBe(true);
    expect(Number.isNaN(parsePresenceTimestamp_("not a date"))).toBe(true);
  });
});

describe("summarizeWorkDay_", () => {
  it("sums a single arrive/leave pair", () => {
    const summary = summarizeWorkDay_(
      [ev("2026-11-21T09:00:00+09:00", "arrive"), ev("2026-11-21T18:30:00+09:00", "leave")],
      "2026-11-21",
    );
    expect(summary).toEqual({
      status: "確定",
      netMinutes: 570,
      firstArrive: "09:00",
      lastLeave: "18:30",
      pairCount: 1,
      segments: [{ startMinute: 540, endMinute: 1110, state: "work" }],
    });
  });

  it("sums split stays instead of bridging a midday break", () => {
    const summary = summarizeWorkDay_(
      [
        ev("2026-11-21T09:00:00+09:00", "arrive"),
        ev("2026-11-21T12:00:00+09:00", "leave"),
        ev("2026-11-21T13:00:00+09:00", "arrive"),
        ev("2026-11-21T18:00:00+09:00", "leave"),
      ],
      "2026-11-21",
    );
    expect(summary.status).toBe("確定");
    expect(summary.netMinutes).toBe(480);
    expect(summary.pairCount).toBe(2);
  });

  it("ignores a duplicate arrive and a duplicate leave", () => {
    const summary = summarizeWorkDay_(
      [
        ev("2026-11-21T09:00:00+09:00", "arrive"),
        ev("2026-11-21T09:05:00+09:00", "arrive"),
        ev("2026-11-21T12:00:00+09:00", "leave"),
        ev("2026-11-21T12:30:00+09:00", "leave"),
        ev("2026-11-21T13:00:00+09:00", "arrive"),
        ev("2026-11-21T18:00:00+09:00", "leave"),
      ],
      "2026-11-21",
    );
    expect(summary.status).toBe("確定");
    expect(summary.netMinutes).toBe(180 + 300);
    expect(summary.pairCount).toBe(2);
    expect(summary.firstArrive).toBe("09:00");
    expect(summary.lastLeave).toBe("18:00");
  });

  it("marks an unclosed arrive as 未確定 with no minutes", () => {
    const summary = summarizeWorkDay_([ev("2026-11-21T09:00:00+09:00", "arrive")], "2026-11-21");
    expect(summary.status).toBe("未確定");
    expect(summary.netMinutes).toBeNull();
    expect(summary.firstArrive).toBe("09:00");
  });

  it("marks a stray leave as 未確定", () => {
    const summary = summarizeWorkDay_([ev("2026-11-21T18:00:00+09:00", "leave")], "2026-11-21");
    expect(summary.status).toBe("未確定");
    expect(summary.netMinutes).toBeNull();
  });

  it("reports 記録なし when there are no work events", () => {
    expect(summarizeWorkDay_([], "2026-11-21").status).toBe("記録なし");
    expect(summarizeWorkDay_([ev("2026-11-21T08:00:00+09:00", "arrive", "home")], "2026-11-21").status).toBe(
      "記録なし",
    );
  });

  it("ignores events from other days", () => {
    const summary = summarizeWorkDay_(
      [ev("2026-11-20T09:00:00+09:00", "arrive"), ev("2026-11-20T18:00:00+09:00", "leave")],
      "2026-11-21",
    );
    expect(summary.status).toBe("記録なし");
  });
});

describe("summarizeWorkDay_ person filter", () => {
  it("ignores other people's work events", () => {
    const events = [
      ev("2026-11-24T09:00:00+09:00", "arrive"),
      { timestamp: "2026-11-24T10:00:00+09:00", person: "aojiru.pudding", place: "work", action: "arrive" },
      { timestamp: "2026-11-24T11:00:00+09:00", person: "aojiru.pudding", place: "work", action: "leave" },
      ev("2026-11-24T18:00:00+09:00", "leave"),
    ];
    const s = summarizeWorkDay_(events, "2026-11-24");
    expect(s.status).toBe("確定");
    expect(s.netMinutes).toBe(540);
    expect(s.pairCount).toBe(1);
  });
});

describe("summarizeWorkDay_ segments", () => {
  it("returns one segment per closed stay, split at a midday break", () => {
    const summary = summarizeWorkDay_(
      [
        ev("2026-11-21T09:00:00+09:00", "arrive"),
        ev("2026-11-21T12:00:00+09:00", "leave"),
        ev("2026-11-21T13:00:00+09:00", "arrive"),
        ev("2026-11-21T18:00:00+09:00", "leave"),
      ],
      "2026-11-21",
    );
    expect(summary.segments).toEqual([
      { startMinute: 540, endMinute: 720, state: "work" },
      { startMinute: 780, endMinute: 1080, state: "work" },
    ]);
  });

  it("keeps closed pairs as segments even when a later arrive is unclosed", () => {
    const summary = summarizeWorkDay_(
      [
        ev("2026-11-21T09:00:00+09:00", "arrive"),
        ev("2026-11-21T12:00:00+09:00", "leave"),
        ev("2026-11-21T13:00:00+09:00", "arrive"),
      ],
      "2026-11-21",
    );
    expect(summary.status).toBe("未確定");
    expect(summary.segments).toEqual([{ startMinute: 540, endMinute: 720, state: "work" }]);
  });

  it("starts the segment at the first arrive when a duplicate arrive follows", () => {
    const summary = summarizeWorkDay_(
      [
        ev("2026-11-21T09:00:00+09:00", "arrive"),
        ev("2026-11-21T09:05:00+09:00", "arrive"),
        ev("2026-11-21T18:00:00+09:00", "leave"),
      ],
      "2026-11-21",
    );
    expect(summary.segments).toEqual([{ startMinute: 540, endMinute: 1080, state: "work" }]);
  });

  it("returns no segments when there are no work events", () => {
    expect(summarizeWorkDay_([], "2026-11-21").segments).toEqual([]);
  });
});

describe("summarizeWorkDay_ flicker tolerance", () => {
  it("merges a stay split by a seconds-long flicker pair", () => {
    const events = [
      ev("2026-11-26T09:00:00+09:00", "arrive", "work"),
      ev("2026-11-26T12:00:00+09:00", "leave", "work"),
      ev("2026-11-26T12:00:30+09:00", "arrive", "work"),
      ev("2026-11-26T18:00:00+09:00", "leave", "work"),
    ];
    const s = summarizeWorkDay_(events, "2026-11-26");
    expect(s.status).toBe("確定");
    expect(s.netMinutes).toBe(540);
    expect(s.pairCount).toBe(1);
    expect(s.segments).toEqual([{ startMinute: 540, endMinute: 1080, state: "work" }]);
  });
});

describe.each([
  ["work_segments", "work", replaceWorkSegments_],
  ["home_segments", "home", replaceHomeSegments_],
])("%s row replacement", (sheetName, state, replaceSegments) => {
  const date = "2026-11-21";
  const now = "2026-11-22T01:02:03.000Z";
  const header = ["date", "start_minute", "end_minute", "state", "minutes", "source", "updated_at"];
  const earlier = ["2026-11-20", 60, 120, state, 60, "presence-gas", now];
  const later = ["2026-11-22", 90, 120, state, 30, "presence-gas", now];

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.useRealTimers();
  });

  it.each([true, false])("rewrites only the target day (append=%s)", (append) => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(now));
    const rows = [
      header,
      [date + "T00:00:00", 100, 110],
      earlier,
      [new Date("2026-11-20T15:00:00Z"), 120, 130],
      later,
    ];
    const sheet = {
      getDataRange: () => ({ getValues: () => rows.map((row) => [...row]) }),
      deleteRow: vi.fn((rowNumber) => rows.splice(rowNumber - 1, 1)),
      appendRow: vi.fn((row) => rows.push(row)),
    };
    const getSheetByName = vi.fn((name) => name === sheetName ? sheet : null);
    vi.stubGlobal("SpreadsheetApp", { openById: () => ({ getSheetByName }) });
    const segments = [
      { startMinute: 540, endMinute: 570, state },
      { startMinute: 600, endMinute: 660, state },
    ];
    replaceSegments(date, append ? segments : []);
    expect(getSheetByName.mock.calls).toEqual([[sheetName]]);
    expect(sheet.deleteRow.mock.calls).toEqual([[4], [2]]);
    const added = append ? [
      [date, 540, 570, state, 30, "presence-gas", now],
      [date, 600, 660, state, 60, "presence-gas", now],
    ] : [];
    expect(sheet.appendRow.mock.calls).toEqual(added.map((row) => [row]));
    expect(rows).toEqual([header, earlier, later, ...added]);
  });
});
