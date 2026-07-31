import {
  addKeepLinesOnlyToAllParagraphs,
  PaginatedParagraph,
  validateKeepLinesPagination,
} from "./paragraph-format";

/* global describe, test, expect */

describe("Keep Paragraphs Intact keepLines-only behavior", () => {
  test("adds keepLines to every paragraph throughout the document body", () => {
    const original =
      "<w:body><w:p><w:r><w:t>First</w:t></w:r></w:p>" +
      "<w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body>";
    const update = addKeepLinesOnlyToAllParagraphs(original);

    expect(update.result).toEqual({
      paragraphsFound: 2,
      paragraphsChanged: 2,
      paragraphsAlreadyFormatted: 0,
    });
    expect(update.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(2);
  });

  test("creates pPr as the first paragraph child when missing", () => {
    const update = addKeepLinesOnlyToAllParagraphs(
      "<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body>"
    );
    expect(update.ooxml).toContain(
      "<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>"
    );
  });

  test("preserves existing paragraph properties", () => {
    const original =
      '<w:body><w:p><w:pPr><w:pStyle w:val="BodyText"/><w:ind w:left="720"/>' +
      '<w:spacing w:after="120"/><w:jc w:val="both"/></w:pPr>' +
      "<w:r><w:t>Exact body</w:t></w:r></w:p></w:body>";
    const update = addKeepLinesOnlyToAllParagraphs(original);
    expect(update.ooxml.replace("<w:keepLines/>", "")).toBe(original);
  });

  test("does not duplicate existing keepLines", () => {
    const original =
      "<w:body><w:p><w:pPr><w:keepLines/></w:pPr>" +
      "<w:r><w:t>Already configured</w:t></w:r></w:p></w:body>";
    const update = addKeepLinesOnlyToAllParagraphs(original);
    expect(update.ooxml.match(/<w:keepLines\/>/g)).toHaveLength(1);
    expect(update.result.paragraphsAlreadyFormatted).toBe(1);
  });

  test("does not add keepNext to lettered headings", () => {
    const original =
      "<w:body><w:p><w:r><w:t>(a) Initial Recognition</w:t></w:r></w:p>" +
      "<w:p><w:r><w:t>Body text</w:t></w:r></w:p></w:body>";
    const update = addKeepLinesOnlyToAllParagraphs(original);
    expect(update.ooxml).not.toContain("<w:keepNext");
  });

  test("does not add keepNext to numbered headings", () => {
    const update = addKeepLinesOnlyToAllParagraphs(
      "<w:body><w:p><w:r><w:t>5.2 Impairment</w:t></w:r></w:p></w:body>"
    );
    expect(update.ooxml).not.toContain("<w:keepNext");
  });

  test("does not create or change CONT'D headings", () => {
    const original =
      "<w:body><w:p><w:r><w:t>5.3 Financial Instruments (Cont'd)</w:t></w:r></w:p></w:body>";
    const update = addKeepLinesOnlyToAllParagraphs(original);
    expect(update.ooxml.replace("<w:pPr><w:keepLines/></w:pPr>", "")).toBe(original);
    expect(update.ooxml.match(/Cont'd/g)).toHaveLength(1);
  });

  test("does not add manual or paragraph page breaks", () => {
    const update = addKeepLinesOnlyToAllParagraphs(
      "<w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p></w:body>"
    );
    expect(update.ooxml).not.toContain("<w:br");
    expect(update.ooxml).not.toContain("pageBreakBefore");
  });

  test("preserves text, runs, hyperlinks, and bookmarks", () => {
    const original =
      '<w:body><w:p><w:bookmarkStart w:id="0" w:name="Exact"/>' +
      '<w:hyperlink w:anchor="Target"><w:r><w:rPr><w:rFonts w:ascii="Aptos"/>' +
      "<w:b/><w:i/></w:rPr><w:t>Exact linked text</w:t></w:r></w:hyperlink>" +
      '<w:bookmarkEnd w:id="0"/></w:p></w:body>';
    const update = addKeepLinesOnlyToAllParagraphs(original);
    expect(update.ooxml.replace("<w:pPr><w:keepLines/></w:pPr>", "")).toBe(original);
  });

  test("formats paragraphs inside tables without changing table structure", () => {
    const original =
      "<w:body><w:tbl><w:tr><w:tc><w:tcPr/>" +
      "<w:p><w:r><w:t>Table paragraph</w:t></w:r></w:p>" +
      "</w:tc></w:tr></w:tbl></w:body>";
    const update = addKeepLinesOnlyToAllParagraphs(original);
    expect(update.ooxml.replace("<w:pPr><w:keepLines/></w:pPr>", "")).toBe(original);
  });
});

describe("Keep Paragraphs Intact repeated pagination validation", () => {
  const item = (text: string, pageCount = 1, properties = ""): PaginatedParagraph => ({
    text,
    pageCount,
    ooxml: `<w:p>${properties}<w:r><w:t>${text}</w:t></w:r></w:p>`,
  });

  const layoutsAfterReflow = (): PaginatedParagraph[][] => [
    [item("Paragraph A", 2), item("Paragraph B")],
    [item("Paragraph A", 1, "<w:pPr><w:keepLines/></w:pPr>"), item("Paragraph B", 2)],
    [
      item("Paragraph A", 1, "<w:pPr><w:keepLines/></w:pPr>"),
      item("Paragraph B", 1, "<w:pPr><w:keepLines/></w:pPr>"),
    ],
  ];

  const runReflow = async () => {
    const layouts = layoutsAfterReflow();
    let layoutIndex = 0;
    let scans = 0;
    let applications = 0;
    const result = await validateKeepLinesPagination(2, async () => {
      const snapshot = layouts[layoutIndex].map((paragraph) => ({ ...paragraph }));
      scans += 1;
      return {
        paragraphs: snapshot,
        applyParagraph: async () => {
          applications += 1;
          layoutIndex = Math.min(layoutIndex + 1, layouts.length - 1);
        },
        settlePagination: async () => undefined,
      };
    });
    return { result, scans, applications };
  };

  test("keeps a paragraph that is split across two pages intact", async () => {
    const { result } = await runReflow();
    expect(result.splitParagraphsFixed).toBe(2);
  });

  test("handles a later paragraph becoming split after the first moves", async () => {
    const { applications } = await runReflow();
    expect(applications).toBe(2);
  });

  test("detects the newly split paragraph on the next pass", async () => {
    const { result } = await runReflow();
    expect(result.paginationPasses).toBe(4);
  });

  test("rescans the document after every repagination", async () => {
    const { scans, applications } = await runReflow();
    expect(scans).toBe(applications + 2);
  });

  test("does not reuse page-position data from an earlier layout", async () => {
    const layouts = layoutsAfterReflow();
    const observedSecondParagraphPages: number[] = [];
    let layoutIndex = 0;
    await validateKeepLinesPagination(2, async () => {
      const snapshot = layouts[layoutIndex].map((paragraph) => ({ ...paragraph }));
      observedSecondParagraphPages.push(snapshot[1].pageCount);
      return {
        paragraphs: snapshot,
        applyParagraph: async () => {
          layoutIndex += 1;
        },
        settlePagination: async () => undefined,
      };
    });
    expect(observedSecondParagraphPages).toEqual([1, 2, 1, 1]);
  });

  test("stops when a complete scan finds no fixable split paragraph", async () => {
    let applications = 0;
    const result = await validateKeepLinesPagination(2, async () => ({
      paragraphs: [item("Complete A"), item("Complete B")],
      applyParagraph: async () => {
        applications += 1;
      },
      settlePagination: async () => undefined,
    }));
    expect(result.paginationPasses).toBe(2);
    expect(applications).toBe(0);
  });

  test("skips and reports a paragraph that remains split with keepLines", async () => {
    const result = await validateKeepLinesPagination(1, async () => ({
      paragraphs: [item("Overlong paragraph", 2, "<w:pPr><w:keepLines/></w:pPr>")],
      applyParagraph: async () => undefined,
      settlePagination: async () => undefined,
    }));
    expect(result.unfixableParagraphs).toBe(1);
    expect(result.splitParagraphsFixed).toBe(0);
    expect(result.paginationPasses).toBeLessThanOrEqual(3);
  });

  test("eventually fixes multiple downstream paragraphs introduced by reflow", async () => {
    const layouts: PaginatedParagraph[][] = [
      [item("A", 2), item("B"), item("C")],
      [item("A", 1, "<w:pPr><w:keepLines/></w:pPr>"), item("B", 2), item("C")],
      [
        item("A", 1, "<w:pPr><w:keepLines/></w:pPr>"),
        item("B", 1, "<w:pPr><w:keepLines/></w:pPr>"),
        item("C", 2),
      ],
      [
        item("A", 1, "<w:pPr><w:keepLines/></w:pPr>"),
        item("B", 1, "<w:pPr><w:keepLines/></w:pPr>"),
        item("C", 1, "<w:pPr><w:keepLines/></w:pPr>"),
      ],
    ];
    let layoutIndex = 0;
    const result = await validateKeepLinesPagination(3, async () => ({
      paragraphs: layouts[layoutIndex].map((paragraph) => ({ ...paragraph })),
      applyParagraph: async () => {
        layoutIndex += 1;
      },
      settlePagination: async () => undefined,
    }));
    expect(result.splitParagraphsFixed).toBe(3);
    expect(result.paginationPasses).toBe(5);
  });
});
