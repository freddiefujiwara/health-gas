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
};
global.Logger = {
  log: vi.fn(),
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
