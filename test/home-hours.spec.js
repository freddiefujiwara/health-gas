import { describe, expect, it } from "vitest";
import { summarizeHomeDay_ } from "../src/Code.js";

const ev = (timestamp, action, place = "home") => ({
  timestamp,
  person: "freddiefujiwara",
  place,
  action,
});

describe("summarizeHomeDay_", () => {
  it("clips an overnight stay at the day boundaries", () => {
    const events = [
      ev("2026-11-23T19:00:00+09:00", "arrive"),
      ev("2026-11-24T04:55:00+09:00", "leave"),
      ev("2026-11-24T19:30:00+09:00", "arrive"),
    ];
    expect(summarizeHomeDay_(events, "2026-11-24")).toEqual({
      status: "確定",
      homeMinutes: 295 + 270,
      firstLeave: "04:55",
      lastArrive: "19:30",
      outingCount: 1,
    });
  });

  it("counts a full day at home as 1440 minutes", () => {
    const events = [ev("2026-11-23T19:00:00+09:00", "arrive")];
    const s = summarizeHomeDay_(events, "2026-11-24");
    expect(s.status).toBe("確定");
    expect(s.homeMinutes).toBe(1440);
    expect(s.outingCount).toBe(0);
  });

  it("counts a full day away as 0 minutes", () => {
    const events = [
      ev("2026-11-23T08:00:00+09:00", "arrive"),
      ev("2026-11-23T09:00:00+09:00", "leave"),
    ];
    const s = summarizeHomeDay_(events, "2026-11-24");
    expect(s.status).toBe("確定");
    expect(s.homeMinutes).toBe(0);
  });

  it("sums multiple outings", () => {
    const events = [
      ev("2026-11-23T20:00:00+09:00", "arrive"),
      ev("2026-11-24T09:00:00+09:00", "leave"),
      ev("2026-11-24T12:00:00+09:00", "arrive"),
      ev("2026-11-24T13:00:00+09:00", "leave"),
      ev("2026-11-24T23:00:00+09:00", "arrive"),
    ];
    const s = summarizeHomeDay_(events, "2026-11-24");
    expect(s.status).toBe("確定");
    expect(s.homeMinutes).toBe(540 + 60 + 60);
    expect(s.outingCount).toBe(2);
  });

  it("is 記録なし when there are no home events at all", () => {
    const s = summarizeHomeDay_([ev("2026-11-24T09:00:00+09:00", "arrive", "work")], "2026-11-24");
    expect(s.status).toBe("記録なし");
    expect(s.homeMinutes).toBeNull();
  });

  it("is 未確定 when the starting state is unknown", () => {
    const events = [ev("2026-11-24T10:00:00+09:00", "arrive")];
    const s = summarizeHomeDay_(events, "2026-11-24");
    expect(s.status).toBe("未確定");
    expect(s.homeMinutes).toBeNull();
  });

  it("is 未確定 on a duplicate arrive", () => {
    const events = [
      ev("2026-11-23T19:00:00+09:00", "arrive"),
      ev("2026-11-24T10:00:00+09:00", "arrive"),
    ];
    expect(summarizeHomeDay_(events, "2026-11-24").status).toBe("未確定");
  });
});

describe("summarizeHomeDay_ person filter", () => {
  it("ignores other people's home events", () => {
    const events = [
      ev("2026-11-23T19:00:00+09:00", "arrive"),
      // The daughter's school-derived home events must not leak in.
      { timestamp: "2026-11-24T07:40:00+09:00", person: "aojiru.pudding", place: "home", action: "leave" },
      { timestamp: "2026-11-24T16:20:00+09:00", person: "aojiru.pudding", place: "home", action: "arrive" },
      ev("2026-11-24T19:30:00+09:00", "leave"),
    ];
    const s = summarizeHomeDay_(events, "2026-11-24");
    expect(s.status).toBe("確定");
    expect(s.homeMinutes).toBe(19 * 60 + 30);
    expect(s.firstLeave).toBe("19:30");
    expect(s.outingCount).toBe(1);
  });
});
