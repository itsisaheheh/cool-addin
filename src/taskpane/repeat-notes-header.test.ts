import { locateNotesSection, pageHasNotesHeader } from "./repeat-notes-header";

/* global describe, expect, test */

const paragraph = (
  text: string,
  page: number,
  listPrefix = ""
): { text: string; listPrefix: string; pages: number[]; startPage: number } => ({
  text,
  listPrefix,
  pages: [page],
  startPage: page,
});

describe("Repeat Notes Header section detection", () => {
  const report = [
    paragraph("INDEPENDENT AUDITORS' REPORT", 1),
    paragraph("STATEMENT OF FINANCIAL POSITION", 2),
    paragraph("NOTES TO THE FINANCIAL STATEMENTS", 3),
    paragraph("For the financial year ended 31 December 2024", 3),
    paragraph("GENERAL INFORMATION", 3, "1."),
    paragraph("SIGNIFICANT ACCOUNTING POLICIES", 4, "5."),
    paragraph("Financial Instruments", 5, "5.2"),
    paragraph("Continued note content", 6),
    paragraph("APPENDIX A", 7),
    paragraph("Appendix content", 8),
  ];

  test("uses the original year-end title and numbered Notes headings to locate Notes pages", () => {
    const location = locateNotesSection(report);

    expect(location).toMatchObject({
      titleIndex: 2,
      yearEndIndex: 3,
      firstNumberedHeadingIndex: 4,
    });
    expect(location?.pages).toEqual([3, 4, 5, 6]);
    expect(report[location?.yearEndIndex ?? -1].text).toBe(
      "For the financial year ended 31 December 2024"
    );
  });

  test("excludes reports, primary statements, and appendix pages", () => {
    const pages = locateNotesSection(report)?.pages ?? [];

    expect(pages).not.toContain(1);
    expect(pages).not.toContain(2);
    expect(pages).not.toContain(7);
    expect(pages).not.toContain(8);
  });

  test("recognizes an existing Notes title as a page-level duplicate", () => {
    expect(pageHasNotesHeader(report, 3)).toBe(true);
    expect(pageHasNotesHeader(report, 4)).toBe(false);
  });

  test("requires the Notes title, year-end line, and a numbered heading", () => {
    expect(locateNotesSection(report.slice(0, 4))).toBeNull();
    expect(locateNotesSection(report.slice(4))).toBeNull();
  });
});
