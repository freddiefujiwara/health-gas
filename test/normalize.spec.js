import { describe, expect, it } from "vitest";
import { collectNumericFixes_, normalizedNumberValue_ } from "../src/Code.js";

describe("normalizedNumberValue_", () => {
  it("converts integer and decimal literals", () => {
    expect(normalizedNumberValue_("88.9")).toBe(88.9);
    expect(normalizedNumberValue_("-841")).toBe(-841);
    expect(normalizedNumberValue_(" 120 ")).toBe(120);
    expect(normalizedNumberValue_("0.67")).toBe(0.67);
  });

  it("leaves non-numeric strings and non-strings alone", () => {
    expect(normalizedNumberValue_("")).toBeNull();
    expect(normalizedNumberValue_("陰性")).toBeNull();
    expect(normalizedNumberValue_("1,234")).toBeNull();
    expect(normalizedNumberValue_("2026-11-21")).toBeNull();
    expect(normalizedNumberValue_("13:44")).toBeNull();
    expect(normalizedNumberValue_(88.9)).toBeNull();
    expect(normalizedNumberValue_(null)).toBeNull();
  });
});

describe("collectNumericFixes_", () => {
  it("finds only the listed columns and only string cells", () => {
    const values = [
      ["date", "weight_kg", "source", "bmi"],
      ["2026-11-21", "88.9", "TANITA", "31.5"],
      ["2026-10-17", 91.5, "TANITA", "32.4"],
      ["2026-09-26", "93.5", "TANITA", 33.1],
    ];
    expect(collectNumericFixes_(values, ["weight_kg", "bmi"])).toEqual([
      { row: 1, col: 1, value: 88.9 },
      { row: 1, col: 3, value: 31.5 },
      { row: 2, col: 3, value: 32.4 },
      { row: 3, col: 1, value: 93.5 },
    ]);
  });

  it("tolerates missing columns and short grids", () => {
    expect(collectNumericFixes_([["date"]], ["weight_kg"])).toEqual([]);
    expect(collectNumericFixes_([], ["weight_kg"])).toEqual([]);
    expect(
      collectNumericFixes_([["date", "x"], ["2026-11-21", "5"]], ["weight_kg"]),
    ).toEqual([]);
  });
});
