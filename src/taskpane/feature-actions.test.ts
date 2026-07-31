jest.mock("./word", () => ({
  checkDocumentIssues: jest.fn(),
  keepAllParagraphsOnOnePage: jest.fn(),
  addContdHeadings: jest.fn(),
}));

import { addContdHeadings, checkDocumentIssues, keepAllParagraphsOnOnePage } from "./word";
import {
  runAddContdHeadingsOnly,
  runCheckDocumentOnly,
  runKeepParagraphsIntactOnly,
} from "./feature-actions";

/* global describe, test, expect, beforeEach, jest */

const mockedCheck = checkDocumentIssues as jest.MockedFunction<typeof checkDocumentIssues>;
const mockedKeep = keepAllParagraphsOnOnePage as jest.MockedFunction<
  typeof keepAllParagraphsOnOnePage
>;
const mockedAddContd = addContdHeadings as jest.MockedFunction<typeof addContdHeadings>;

describe("independent feature actions", () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test("Keep Paragraphs Intact does not call the CONT'D insertion function", async () => {
    mockedKeep.mockResolvedValue({
      paragraphsFound: 1,
      paragraphsChanged: 0,
      paragraphsAlreadyFormatted: 0,
      splitParagraphsFixed: 0,
      orphanHeadingsFixed: 0,
      paginationPasses: 2,
      unfixableParagraphs: 0,
    });

    await runKeepParagraphsIntactOnly();

    expect(mockedKeep).toHaveBeenCalledTimes(1);
    expect(mockedAddContd).not.toHaveBeenCalled();
    expect(mockedCheck).not.toHaveBeenCalled();
  });

  test("Check Document does not call the CONT'D insertion function", async () => {
    mockedCheck.mockResolvedValue({
      numberedHeadings: [],
    });

    await runCheckDocumentOnly();

    expect(mockedCheck).toHaveBeenCalledTimes(1);
    expect(mockedAddContd).not.toHaveBeenCalled();
    expect(mockedKeep).not.toHaveBeenCalled();
  });

  test("Add CONT'D Headings calls only the CONT'D insertion function", async () => {
    mockedAddContd.mockResolvedValue({
      continuingSectionsFound: 1,
      continuationPagesFound: 1,
      headingsInserted: 1,
      duplicatesSkipped: 0,
      limitationMessage: "Inserted.",
    });

    const result = await runAddContdHeadingsOnly();

    expect(mockedAddContd).toHaveBeenCalledTimes(1);
    expect(mockedCheck).not.toHaveBeenCalled();
    expect(mockedKeep).not.toHaveBeenCalled();
    expect(result.headingsInserted).toBe(1);
  });

  test("Keep and Check actions cannot return inserted CONT'D text", async () => {
    mockedKeep.mockResolvedValue({
      paragraphsFound: 1,
      paragraphsChanged: 0,
      paragraphsAlreadyFormatted: 0,
      splitParagraphsFixed: 0,
      orphanHeadingsFixed: 0,
      paginationPasses: 2,
      unfixableParagraphs: 0,
    });
    mockedCheck.mockResolvedValue({
      numberedHeadings: [
        {
          key: "5.2",
          title: "Property",
        },
      ],
    });

    const results = [await runKeepParagraphsIntactOnly(), await runCheckDocumentOnly()];
    const serialized = JSON.stringify(results);

    expect(serialized).not.toMatch(/\(Cont['’]d\)|CONT'D/);
    expect(mockedAddContd).not.toHaveBeenCalled();
  });

  test("Add CONT'D Headings preserves insertion and duplicate-prevention results", async () => {
    mockedAddContd.mockResolvedValue({
      continuingSectionsFound: 2,
      continuationPagesFound: 2,
      headingsInserted: 1,
      duplicatesSkipped: 1,
      limitationMessage: "Existing duplicate skipped.",
    });

    const result = await runAddContdHeadingsOnly();

    expect(result.headingsInserted).toBe(1);
    expect(result.duplicatesSkipped).toBe(1);
  });
});
