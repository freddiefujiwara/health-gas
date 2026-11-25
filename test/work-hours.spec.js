import { describe, expect, it } from "vitest";
import { parsePresenceTimestamp_, summarizeWorkDay_ } from "../src/Code.js";

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
