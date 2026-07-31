import { findNumberedHeadingsInNotes, inspectNumberedHeadingsInNotes } from "./check-document";

/* global describe, expect, test */

describe("Check Document Notes-only numbered headings", () => {
  test("ignores numbered headings before the Notes section", () => {
    const result = findNumberedHeadingsInNotes([
      { text: "1 Directors' Report", listPrefix: "" },
      { text: "2 Statement of Financial Position", listPrefix: "" },
      { text: "NOTES TO THE FINANCIAL STATEMENTS", listPrefix: "" },
      { text: "4.1 Basis of Preparation", listPrefix: "" },
    ]);

    expect(result).toEqual([{ key: "4.1", title: "Basis of Preparation" }]);
  });

  test("detects all generic numbered headings after the Notes heading in order", () => {
    const result = findNumberedHeadingsInNotes([
      { text: "  Notes   to the Financial Statements  ", listPrefix: "" },
      { text: "1 General Information", listPrefix: "" },
      { text: "4.6 Revenue Recognition", listPrefix: "" },
      { text: "7.2.1 Sale of Goods", listPrefix: "" },
      { text: "10.1 Bank Overdrafts", listPrefix: "" },
    ]);

    expect(result.map(({ key }) => key)).toEqual(["1", "4.6", "7.2.1", "10.1"]);
  });

  test("supports Word list numbering within Notes", () => {
    const result = findNumberedHeadingsInNotes([
      { text: "NOTES TO THE FINANCIAL STATEMENTS", listPrefix: "" },
      { text: "Inventories", listPrefix: "12." },
    ]);

    expect(result).toEqual([{ key: "12", title: "Inventories" }]);
  });

  test("returns no headings when the Notes heading is absent", () => {
    expect(
      findNumberedHeadingsInNotes([
        { text: "1 Directors' Report", listPrefix: "" },
        { text: "10 Statement of Cash Flows", listPrefix: "" },
      ])
    ).toEqual([]);
  });

  test("recognizes an exact visible Notes heading containing Word control characters", () => {
    const result = inspectNumberedHeadingsInNotes([
      { text: "NOTES TO THE FINANCIAL STATEMENTS\r\u0007", listPrefix: "" },
      { text: "General Information\r\u0007", listPrefix: "１." },
    ]);

    expect(result.headings).toEqual([{ key: "1", title: "General Information" }]);
    expect(result.diagnostics).toMatchObject({
      notesHeadingFound: true,
      notesHeadingIndex: 0,
      paragraphsScannedAfterNotes: 1,
    });
  });

  test("reports the first 20 non-empty paragraphs and rejection reasons after Notes", () => {
    const result = inspectNumberedHeadingsInNotes([
      { text: "Before Notes", listPrefix: "" },
      { text: "NOTES TO THE FINANCIAL STATEMENTS", listPrefix: "" },
      { text: "General Information", listPrefix: "" },
      { text: "Inventories", listPrefix: "12." },
    ]);

    expect(result.diagnostics.firstNonEmptyParagraphs).toEqual([
      {
        paragraphIndex: 2,
        text: "General Information",
        listPrefix: "",
        result: "Rejected: no numeric prefix in paragraph.text or Word listString.",
      },
      {
        paragraphIndex: 3,
        text: "Inventories",
        listPrefix: "12.",
        result: "Accepted as section 12.",
      },
    ]);
  });
});
