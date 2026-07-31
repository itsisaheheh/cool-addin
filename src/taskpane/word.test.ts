import {
  addKeepLinesToAllParagraphs,
  disableMoveToNextPage,
  enableKeepTogether,
  removeKeepLinesFromAllParagraphs,
} from "./paragraph-format";
import { continuationText, parseNumericHeading } from "./continuation-format";

/* global describe, test, expect */

describe("Move Paragraph to Next Page OOXML", () => {
  const paragraph = (pageCount: number, properties = "", runs = "<w:r><w:t>Text</w:t></w:r>") => ({
    ooxml: `<w:p${properties}><w:pPr><w:jc w:val="center"/></w:pPr>${runs}</w:p>`,
    pageCount,
  });

  test("adds Page break before and Keep lines together to a split paragraph", () => {
    const update = enableKeepTogether([paragraph(2)]);

    expect(update).toMatchObject({ paragraphsUpdated: 1, splitParagraphsFound: 1 });
    expect(update.paragraphs[0]).toContain("<w:pageBreakBefore/>");
    expect(update.paragraphs[0]).toContain("<w:keepLines/>");
  });

  test("adds Page break before when the paragraph already fits on one page", () => {
    const update = enableKeepTogether([paragraph(1)]);

    expect(update.splitParagraphsFound).toBe(0);
    expect(update.paragraphs[0]).toContain("<w:pageBreakBefore/>");
    expect(update.paragraphs[0]).not.toContain("<w:br");
  });

  test("applies Page break before to every selected paragraph", () => {
    const update = enableKeepTogether([paragraph(1), paragraph(2), paragraph(1)]);

    expect(update).toMatchObject({ paragraphsUpdated: 3, splitParagraphsFound: 1 });
    expect(update.paragraphs.every((ooxml) => ooxml.includes("<w:pageBreakBefore/>"))).toBe(true);
  });

  test("handles a cursor-only selection represented by its containing paragraph", () => {
    const update = enableKeepTogether([paragraph(2)]);

    expect(update.paragraphsUpdated).toBe(1);
    expect(update.paragraphs[0]).toContain("<w:pageBreakBefore/>");
  });

  test("preserves content and existing formatting when adding move formatting", () => {
    const original =
      '<w:p w:rsidR="1234"><w:pPr><w:pStyle w:val="Numbered"/><w:jc w:val="right"/>' +
      '<w:ind w:left="720"/><w:spacing w:before="120" w:after="240"/></w:pPr>' +
      '<w:r><w:rPr><w:rFonts w:ascii="Aptos"/><w:b/></w:rPr><w:t>2.1 Exact text</w:t></w:r></w:p>';

    const updated = enableKeepTogether([{ ooxml: original, pageCount: 2 }]).paragraphs[0];

    expect(updated.replace("<w:pageBreakBefore/>", "").replace("<w:keepLines/>", "")).toBe(
      original
    );
    expect(updated).not.toContain("CONT");
  });

  test("removes only Page break before and retains Keep lines and other formatting", () => {
    const original =
      '<w:p><w:pPr><w:pageBreakBefore/><w:keepLines/><w:numPr><w:ilvl w:val="1"/></w:numPr>' +
      '<w:ind w:left="720"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>Exact text</w:t></w:r></w:p>';

    const update = disableMoveToNextPage([original]);

    expect(update.result).toEqual({ paragraphsUpdated: 1, pageBreaksRemoved: 1 });
    expect(update.paragraphs[0]).not.toContain("pageBreakBefore");
    expect(update.paragraphs[0]).toContain("<w:keepLines/>");
    expect(update.paragraphs[0]).toBe(original.replace("<w:pageBreakBefore/>", ""));
  });
});

describe("continuation heading text", () => {
  test("recognizes a numbered 1.1 heading and appends Cont'd", () => {
    expect(parseNumericHeading("1.1 Scope")).toEqual({ key: "1.1", level: 2 });
    expect(continuationText("1.1 Scope")).toBe("1.1 Scope (Cont'd)");
  });
});

describe("Keep All Paragraphs on One Page OOXML", () => {
  test("formats multiple paragraphs across the whole document", () => {
    const original =
      "<w:body><w:p><w:r><w:t>One</w:t></w:r></w:p>" +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Two</w:t></w:r></w:p></w:body>';

    const update = addKeepLinesToAllParagraphs(original);

    expect(update.result).toEqual({
      paragraphsFound: 2,
      paragraphsChanged: 2,
      paragraphsAlreadyFormatted: 0,
    });
    expect(update.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(2);
  });

  test("adds keepLines to a paragraph that would otherwise split across pages", () => {
    const original = "<w:body><w:p><w:r><w:t>Split paragraph</w:t></w:r></w:p></w:body>";

    const update = addKeepLinesToAllParagraphs(original);

    expect(update.ooxml).toContain("<w:p><w:pPr><w:keepLines/></w:pPr>");
    expect(update.ooxml).not.toContain("pageBreakBefore");
  });

  test("formats paragraphs inside tables", () => {
    const original =
      "<w:body><w:tbl><w:tr><w:tc><w:tcPr/>" +
      "<w:p><w:r><w:t>Table cell</w:t></w:r></w:p>" +
      "</w:tc></w:tr></w:tbl></w:body>";

    const update = addKeepLinesToAllParagraphs(original);

    expect(update.result.paragraphsFound).toBe(1);
    expect(update.ooxml).toContain(
      "<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Table cell</w:t></w:r></w:p>"
    );
  });

  test("preserves numbered and bulleted paragraph properties", () => {
    const original =
      '<w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="4"/></w:numPr>' +
      '<w:pStyle w:val="ListParagraph"/></w:pPr><w:r><w:t>List item</w:t></w:r></w:p></w:body>';

    const update = addKeepLinesToAllParagraphs(original);

    expect(update.ooxml.replace("<w:keepLines/>", "")).toBe(original);
  });

  test("preserves all existing paragraph and content formatting", () => {
    const original =
      '<w:body><w:p w:rsidR="ABCD"><w:pPr><w:pStyle w:val="Heading1"/>' +
      '<w:ind w:left="720"/><w:spacing w:before="120"/><w:jc w:val="right"/></w:pPr>' +
      '<w:bookmarkStart w:id="0" w:name="Bookmark"/><w:hyperlink w:anchor="Target">' +
      '<w:r><w:rPr><w:rFonts w:ascii="Aptos"/><w:b/></w:rPr><w:t>Linked heading</w:t></w:r>' +
      "</w:hyperlink></w:p></w:body>";

    const update = addKeepLinesToAllParagraphs(original);

    expect(update.ooxml.replace("<w:keepLines/>", "")).toBe(original);
  });

  test("does not duplicate an existing keepLines property", () => {
    const original =
      '<w:body><w:p><w:pPr><w:keepLines/><w:jc w:val="left"/></w:pPr>' +
      "<w:r><w:t>Already set</w:t></w:r></w:p></w:body>";

    const update = addKeepLinesToAllParagraphs(original);

    expect(update.result).toEqual({
      paragraphsFound: 1,
      paragraphsChanged: 0,
      paragraphsAlreadyFormatted: 1,
    });
    expect(update.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(1);
  });

  test("removes only keepLines and preserves other paragraph properties", () => {
    const original =
      '<w:body><w:p><w:pPr><w:pStyle w:val="BodyText"/><w:keepLines/>' +
      '<w:pageBreakBefore/><w:ind w:right="360"/></w:pPr>' +
      "<w:r><w:t>Exact text</w:t></w:r></w:p></w:body>";

    const update = removeKeepLinesFromAllParagraphs(original);

    expect(update.result.paragraphsChanged).toBe(1);
    expect(update.ooxml).toBe(original.replace("<w:keepLines/>", ""));
    expect(update.ooxml).toContain("<w:pageBreakBefore/>");
  });

  test("handles a paragraph longer than one page with one idempotent formatting change", () => {
    const longText = "Long paragraph content ".repeat(10000);
    const original = `<w:body><w:p><w:r><w:t>${longText}</w:t></w:r></w:p></w:body>`;

    const firstUpdate = addKeepLinesToAllParagraphs(original);
    const secondUpdate = addKeepLinesToAllParagraphs(firstUpdate.ooxml);

    expect(firstUpdate.result.paragraphsChanged).toBe(1);
    expect(secondUpdate.result.paragraphsChanged).toBe(0);
    expect(secondUpdate.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(1);
    expect(secondUpdate.ooxml).not.toContain("pageBreakBefore");
  });
});
