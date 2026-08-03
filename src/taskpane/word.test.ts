import {
  addKeepNextToParagraphOoxml,
  addKeepLinesToAllParagraphs,
  disableMoveToNextPage,
  enableKeepTogether,
  removeKeepLinesFromAllParagraphs,
} from "./paragraph-format";
import {
  continuationPageEligibility,
  continuationPlacement,
  continuationText,
  isOrphanOriginalHeading,
  MAX_CONTINUATION_PAGINATION_PASSES,
  parseNumericHeading,
} from "./continuation-format";

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

  test("uses the report's uppercase CONT'D suffix for a main heading", () => {
    expect(continuationText("5. SIGNIFICANT ACCOUNTING POLICIES", true)).toBe(
      "5. SIGNIFICANT ACCOUNTING POLICIES (CONT'D)"
    );
  });

  test("keeps parent and subsection continuation headings in display order", () => {
    const hierarchy = [
      continuationText("5. SIGNIFICANT ACCOUNTING POLICIES", true),
      continuationText("5.2 Financial Instruments"),
    ];

    expect(hierarchy).toEqual([
      "5. SIGNIFICANT ACCOUNTING POLICIES (CONT'D)",
      "5.2 Financial Instruments (Cont'd)",
    ]);
  });

  test.each([
    ["1", "1", 1],
    ["1.1 Basis of Preparation", "1.1", 2],
    ["7.2.1 Detailed Policy", "7.2.1", 3],
    ["10 BORROWINGS", "10", 1],
    ["20.3 Related Party Balances", "20.3", 2],
    ["40.1 Other Information", "40.1", 2],
  ])("recognizes generic numbered report heading %s", (text, key, level) => {
    expect(parseNumericHeading(text as string)).toEqual({ key, level });
  });
});

describe("post-CONT'D pagination validation", () => {
  const bodyParagraph =
    '<w:p><w:pPr><w:jc w:val="justified"/></w:pPr><w:r><w:rPr><w:i/></w:rPr>' +
    "<w:t>Complete paragraph text remains in one paragraph.</w:t></w:r></w:p>";

  test("prepares a paragraph that is already split before CONT'D", () => {
    expect(continuationPlacement(true, true)).toBe("prepare-and-repaginate");
    expect(addKeepLinesToAllParagraphs(bodyParagraph).ooxml).toContain("<w:keepLines/>");
  });

  test("keeps a paragraph intact if it would split after CONT'D insertion", () => {
    const once = addKeepLinesToAllParagraphs(bodyParagraph);
    const twice = addKeepLinesToAllParagraphs(once.ooxml);
    expect(once.result.paragraphsChanged).toBe(1);
    expect(twice.result.paragraphsChanged).toBe(0);
  });

  test("places the CONT'D heading before the complete paragraph", () => {
    const heading = addKeepNextToParagraphOoxml(
      addKeepLinesToAllParagraphs(
        `<w:p><w:r><w:t>${continuationText("5.3 Financial Instruments (a)")}</w:t></w:r></w:p>`
      ).ooxml
    );
    const layout = `${heading}${addKeepLinesToAllParagraphs(bodyParagraph).ooxml}`;
    expect(layout.indexOf("(Cont'd)")).toBeLessThan(layout.indexOf("Complete paragraph text"));
    expect(heading).toContain("<w:keepNext/><w:keepLines/>");
  });

  test("inserts before the first continuation item that starts on the page", () => {
    expect(
      continuationPageEligibility({
        sectionStartPage: 10,
        currentPage: 11,
        anchorStartPage: 11,
        anchorIsOriginalHeading: false,
        anchorSpansFromEarlierPage: false,
      })
    ).toBe("insert");
  });

  test("prepares a paragraph spanning into the continuation page before insertion", () => {
    expect(
      continuationPageEligibility({
        sectionStartPage: 10,
        currentPage: 11,
        anchorStartPage: 10,
        anchorIsOriginalHeading: false,
        anchorSpansFromEarlierPage: true,
      })
    ).toBe("prepare");
  });

  test("does not add CONT'D beside an original section heading on the current page", () => {
    expect(
      continuationPageEligibility({
        sectionStartPage: 11,
        currentPage: 11,
        anchorStartPage: 11,
        anchorIsOriginalHeading: true,
        anchorSpansFromEarlierPage: false,
      })
    ).toBe("skip");
  });

  test("recognizes an original heading orphaned from its first content paragraph", () => {
    expect(
      isOrphanOriginalHeading({
        headingStartPage: 1,
        nextContentStartPage: 2,
        nextParagraphIsNumberedHeading: false,
      })
    ).toBe(true);
  });

  test("does not treat a heading followed by same-page content as orphaned", () => {
    expect(
      isOrphanOriginalHeading({
        headingStartPage: 2,
        nextContentStartPage: 2,
        nextParagraphIsNumberedHeading: false,
      })
    ).toBe(false);
  });

  test("does not treat a heading followed only by a new numbered section as orphaned", () => {
    expect(
      isOrphanOriginalHeading({
        headingStartPage: 1,
        nextContentStartPage: 2,
        nextParagraphIsNumberedHeading: true,
      })
    ).toBe(false);
  });

  test("does not insert a CONT'D heading inside paragraph text", () => {
    const updated = addKeepLinesToAllParagraphs(bodyParagraph).ooxml;
    expect(updated).toContain("<w:t>Complete paragraph text remains in one paragraph.</w:t>");
    expect(updated).not.toContain("Cont'd");
  });

  test("supports multiple CONT'D headings without merging them", () => {
    const headings = [continuationText("5.3 First"), continuationText("5.4 Second")];
    expect(headings).toEqual(["5.3 First (Cont'd)", "5.4 Second (Cont'd)"]);
  });

  test("recognizes a paragraph inside a numbered section", () => {
    expect(parseNumericHeading("5.3 Financial Instruments")).toEqual({ key: "5.3", level: 2 });
  });

  test("uses the rendered page start when a paragraph still spans the continuation page", () => {
    expect(continuationPlacement(true, false)).toBe("at-rendered-page-start");
  });

  test("preserves existing body paragraph formatting", () => {
    const updated = addKeepLinesToAllParagraphs(bodyParagraph).ooxml;
    expect(updated.replace("<w:keepLines/>", "")).toBe(bodyParagraph);
  });

  test("does not treat an existing CONT'D heading as a new source heading", () => {
    expect(parseNumericHeading("5.3 Financial Instruments (Cont'd)")).toBeNull();
  });

  test("uses a bounded two-pass workflow and cannot loop indefinitely", () => {
    expect(MAX_CONTINUATION_PAGINATION_PASSES).toBe(2);
    expect(continuationPlacement(false, false)).toBe("before-complete-paragraph");
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

    expect(update.ooxml.replace("<w:keepNext/>", "").replace("<w:keepLines/>", "")).toBe(original);
    expect(update.ooxml).toContain("<w:keepNext/>");
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

describe("Keep Paragraphs Intact heading pairing", () => {
  const documentWith = (heading: string, body = "Complete body paragraph") =>
    `<w:body>${heading}<w:p><w:r><w:t>${body}</w:t></w:r></w:p></w:body>`;
  const paragraph = (text: string, properties = "", runProperties = "") =>
    `<w:p>${properties}<w:r>${runProperties}<w:t>${text}</w:t></w:r></w:p>`;

  test("keeps a numbered note heading with its following paragraph", () => {
    const update = addKeepLinesToAllParagraphs(
      documentWith(paragraph("5.4 Tax assets and liabilities"))
    );

    expect(update.ooxml).toContain(
      "<w:keepNext/><w:keepLines/></w:pPr><w:r><w:t>5.4 Tax assets and liabilities"
    );
    expect(update.ooxml).toContain(
      "<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Complete body paragraph"
    );
  });

  test("keeps a lettered subsection heading with its following paragraph", () => {
    const update = addKeepLinesToAllParagraphs(
      documentWith(paragraph("(f) Recognition of Gains and Losses"))
    );

    expect(update.ooxml).toContain("<w:keepNext/><w:keepLines/>");
    expect(update.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(2);
  });

  test("keeps a non-numbered report heading with its following paragraph", () => {
    const update = addKeepLinesToAllParagraphs(documentWith(paragraph("DIRECTORS’ BENEFITS")));

    expect(update.ooxml).toContain("<w:keepNext/><w:keepLines/>");
  });

  test("does not duplicate an existing keepNext property", () => {
    const heading = paragraph(
      "12. Inventories",
      '<w:pPr><w:keepNext/><w:spacing w:after="120"/></w:pPr>'
    );
    const update = addKeepLinesToAllParagraphs(documentWith(heading));

    expect(update.ooxml.match(/<w:keepNext\/>/g)).toHaveLength(1);
  });

  test("does not duplicate an existing keepLines property", () => {
    const heading = paragraph("20. Cash and Cash Equivalents", "<w:pPr><w:keepLines/></w:pPr>");
    const body = "<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Cash details</w:t></w:r></w:p>";
    const update = addKeepLinesToAllParagraphs(`<w:body>${heading}${body}</w:body>`);

    expect(update.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(2);
  });

  test("keeps heading and body text in separate paragraphs", () => {
    const update = addKeepLinesToAllParagraphs(
      documentWith(paragraph("(a) Initial Recognition and Measurement"))
    );

    expect(update.result.paragraphsFound).toBe(2);
    expect(update.ooxml.match(/<w:p(?:\s[^>]*)?>/g)).toHaveLength(2);
    expect(update.ooxml).toMatch(/Initial Recognition and Measurement<\/w:t><\/w:r><\/w:p><w:p>/);
  });

  test("does not insert a manual page break", () => {
    const update = addKeepLinesToAllParagraphs(documentWith(paragraph("12. Inventories")));

    expect(update.ooxml).not.toContain("<w:br");
  });

  test("does not insert pageBreakBefore", () => {
    const update = addKeepLinesToAllParagraphs(documentWith(paragraph("12. Inventories")));

    expect(update.ooxml).not.toContain("pageBreakBefore");
  });

  test("does not add a CONT'D heading to numbered Notes content", () => {
    const update = addKeepLinesToAllParagraphs(
      documentWith(paragraph("5.4 Tax assets and liabilities"))
    );

    expect(update.ooxml).not.toMatch(/CONT['’]?D|Cont['’]?d/);
    expect(update.ooxml.match(/Tax assets and liabilities/g)).toHaveLength(1);
  });

  test("uses heading style and bold-short-text signals without changing content formatting", () => {
    const heading = paragraph(
      "Accounting policies",
      '<w:pPr><w:pStyle w:val="Heading2"/><w:ind w:left="240"/></w:pPr>',
      "<w:rPr><w:b/><w:i/></w:rPr>"
    );
    const update = addKeepLinesToAllParagraphs(documentWith(heading));

    expect(update.ooxml).toContain('<w:pStyle w:val="Heading2"/>');
    expect(update.ooxml).toContain('<w:ind w:left="240"/>');
    expect(update.ooxml).toContain("<w:rPr><w:b/><w:i/></w:rPr>");
    expect(update.ooxml).toContain("<w:keepNext/>");
  });
});
