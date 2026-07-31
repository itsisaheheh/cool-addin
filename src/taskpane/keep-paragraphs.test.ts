import { formatSplitParagraphChains } from "./paragraph-format";

/* global describe, test, expect */

describe("Keep Paragraphs Intact independence", () => {
  const item = (text: string, pageCount = 1, properties = "", runProperties = "") => ({
    text,
    pageCount,
    ooxml: `<w:p>${properties}<w:r>${runProperties}<w:t>${text}</w:t></w:r></w:p>`,
  });

  test("adds keepLines only to a body paragraph confirmed split across pages", () => {
    const update = formatSplitParagraphChains([item("Split body paragraph.", 2)]);
    expect(update.paragraphs[0]).toContain("<w:keepLines/>");
    expect(update.bodyIndices).toEqual([0]);
  });

  test("leaves an intact body paragraph unchanged", () => {
    const original = item("Intact body paragraph.");
    const update = formatSplitParagraphChains([original]);
    expect(update.paragraphs[0]).toBe(original.ooxml);
    expect(update.changedIndices).toEqual([]);
  });

  test.each([
    "(a) Initial Recognition and Measurement",
    "(b) Subsequent Measurement",
    "(c) Impairment",
    "(f) Recognition of Gains and Losses",
  ])("keeps immediate lettered topic heading %s with split body", (heading) => {
    const update = formatSplitParagraphChains([item(heading), item("Split body paragraph.", 2)]);
    expect(update.paragraphs[0]).toContain("<w:keepNext/>");
    expect(update.paragraphs[1]).toContain("<w:keepLines/>");
    expect(update.headingIndices).toEqual([0]);
  });

  test("does not apply keepNext to a numbered note heading", () => {
    const update = formatSplitParagraphChains([
      item("5.3 Financial Instruments"),
      item("Split body paragraph.", 2),
    ]);
    expect(update.paragraphs[0]).not.toContain("<w:keepNext/>");
    expect(update.paragraphs[1]).toContain("<w:keepLines/>");
  });

  test("does not build a numbered and lettered heading chain", () => {
    const update = formatSplitParagraphChains([
      item("5.3 Financial Instruments"),
      item("(f) Recognition of Gains and Losses"),
      item("Split body paragraph.", 2),
    ]);
    expect(update.paragraphs[0]).not.toContain("<w:keepNext/>");
    expect(update.paragraphs[1]).toContain("<w:keepNext/>");
    expect(update.paragraphs[2]).toContain("<w:keepLines/>");
  });

  test("does not alter an existing numbered CONT'D paragraph", () => {
    const continued = item("5.3 Financial Instruments (Cont'd)");
    const update = formatSplitParagraphChains([continued, item("Split body paragraph.", 2)]);
    expect(update.paragraphs[0]).toBe(continued.ooxml);
    expect(update.paragraphs.join("")).toMatch(/5\.3 Financial Instruments \(Cont'd\)/);
  });

  test("does not insert CONT'D text", () => {
    const update = formatSplitParagraphChains([
      item("(a) Initial Recognition and Measurement"),
      item("Split body paragraph.", 2),
    ]);
    expect(update.paragraphs.join("")).not.toMatch(/CONT['’]?D/i);
  });

  test("does not treat an ordinary bold sentence as a topic heading", () => {
    const update = formatSplitParagraphChains([
      item(
        "This ordinary bold sentence explains the accounting treatment.",
        1,
        "",
        "<w:rPr><w:b/></w:rPr>"
      ),
      item("Split body paragraph.", 2),
    ]);
    expect(update.paragraphs[0]).not.toContain("<w:keepNext/>");
  });

  test("does not duplicate existing keep properties", () => {
    const update = formatSplitParagraphChains([
      item("(a) Topic", 1, "<w:pPr><w:keepNext/></w:pPr>"),
      item("Split body.", 2, "<w:pPr><w:keepLines/></w:pPr>"),
    ]);
    expect(update.paragraphs[0].match(/<w:keepNext\/>/g)).toHaveLength(1);
    expect(update.paragraphs[1].match(/<w:keepLines\/>/g)).toHaveLength(1);
  });

  test("preserves text, runs, and existing paragraph properties", () => {
    const original = item(
      "Exact body text",
      2,
      '<w:pPr><w:ind w:left="720"/><w:spacing w:after="120"/></w:pPr>',
      '<w:rPr><w:rFonts w:ascii="Aptos"/><w:b/><w:i/></w:rPr>'
    );
    const update = formatSplitParagraphChains([original]);
    expect(update.paragraphs[0].replace("<w:keepLines/>", "")).toBe(original.ooxml);
  });

  test("adds no manual or paragraph page break", () => {
    const update = formatSplitParagraphChains([
      item("(f) Recognition of Gains and Losses"),
      item("Split body paragraph.", 2),
    ]);
    expect(update.paragraphs.join("")).not.toContain("<w:br");
    expect(update.paragraphs.join("")).not.toContain("pageBreakBefore");
  });
});
