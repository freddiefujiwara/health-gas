import { describe, expect, it } from "vitest";
import { isTimeOnlyDate_ } from "../src/Code.js";

describe("isTimeOnlyDate_", () => {
  it("detects the Sheets epoch day used for time-only cells", () => {
    expect(isTimeOnlyDate_(new Date(1899, 11, 30, 8, 11))).toBe(true);
    expect(isTimeOnlyDate_(new Date(1899, 11, 30, 0, 0))).toBe(true);
  });

  it("rejects real dates and non-dates", () => {
    expect(isTimeOnlyDate_(new Date(2026, 10, 24, 0, 0))).toBe(false);
    expect(isTimeOnlyDate_(new Date(1899, 11, 31, 0, 0))).toBe(false);
    expect(isTimeOnlyDate_("08:11")).toBe(false);
    expect(isTimeOnlyDate_(null)).toBe(false);
  });
});
