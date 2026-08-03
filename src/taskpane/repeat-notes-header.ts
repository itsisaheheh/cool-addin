import { CONTINUATION_SUFFIX_PATTERN, parseNumericHeading } from "./continuation-format";

/* global Office, Word */

const NOTES_TITLE_PATTERN = /^NOTES TO THE FINANCIAL STATEMENTS$/i;
const NOTES_YEAR_END_PATTERN = /^FOR THE FINANCIAL YEAR ENDED\b/i;
const NOTES_END_PATTERN = /^(?:APPENDI(?:X|CES)\b|SUPPLEMENTARY INFORMATION\b)/i;
const NOTES_HEADER_TAG_PREFIX = "word-notes-header:";

export interface NotesPageParagraph {
  text: string;
  listPrefix?: string;
  pages: number[];
  startPage: number | null;
}

export interface NotesSectionLocation {
  titleIndex: number;
  yearEndIndex: number;
  firstNumberedHeadingIndex: number;
  endIndex: number;
  pages: number[];
}

export interface RepeatNotesHeaderResult {
  notesPagesFound: number;
  headersInserted: number;
  duplicatesSkipped: number;
}

const normalizeWordText = (text: string): string =>
  text
    .normalize("NFKC")
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u001f\u007f-\u009f\u200b-\u200f\u202a-\u202e\u2060\ufeff]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const numericHeading = (paragraph: NotesPageParagraph) => {
  const text = normalizeWordText(paragraph.text);
  const listPrefix = normalizeWordText(paragraph.listPrefix ?? "");
  return parseNumericHeading(text) ?? parseNumericHeading(`${listPrefix} `);
};

export function locateNotesSection(paragraphs: NotesPageParagraph[]): NotesSectionLocation | null {
  const titleIndex = paragraphs.findIndex((paragraph) =>
    NOTES_TITLE_PATTERN.test(normalizeWordText(paragraph.text))
  );
  if (titleIndex < 0) return null;

  const yearEndIndex = paragraphs.findIndex(
    (paragraph, index) => index > titleIndex && normalizeWordText(paragraph.text) !== ""
  );
  if (
    yearEndIndex < 0 ||
    !NOTES_YEAR_END_PATTERN.test(normalizeWordText(paragraphs[yearEndIndex].text))
  ) {
    return null;
  }

  const firstNumberedHeadingIndex = paragraphs.findIndex(
    (paragraph, index) => index > yearEndIndex && numericHeading(paragraph) !== null
  );
  if (firstNumberedHeadingIndex < 0) return null;

  const explicitEndIndex = paragraphs.findIndex(
    (paragraph, index) =>
      index > firstNumberedHeadingIndex && NOTES_END_PATTERN.test(normalizeWordText(paragraph.text))
  );
  const endIndex = explicitEndIndex < 0 ? paragraphs.length : explicitEndIndex;
  const startPage = paragraphs[titleIndex].startPage;
  const endPage = explicitEndIndex < 0 ? null : paragraphs[explicitEndIndex].startPage;
  if (startPage === null) return null;

  const pages = Array.from(
    new Set(
      paragraphs
        .slice(titleIndex, endIndex)
        .flatMap((paragraph) => paragraph.pages)
        .filter((page) => page >= startPage && (endPage === null || page < endPage))
    )
  ).sort((left, right) => left - right);

  return { titleIndex, yearEndIndex, firstNumberedHeadingIndex, endIndex, pages };
}

export function pageHasNotesHeader(paragraphs: NotesPageParagraph[], pageIndex: number): boolean {
  return paragraphs.some(
    (paragraph) =>
      paragraph.pages.includes(pageIndex) &&
      NOTES_TITLE_PATTERN.test(normalizeWordText(paragraph.text))
  );
}

const copyParagraphFormatting = (source: Word.Paragraph, target: Word.Paragraph): void => {
  target.style = source.style;
  target.alignment = source.alignment;
  target.firstLineIndent = source.firstLineIndent;
  target.leftIndent = source.leftIndent;
  target.rightIndent = source.rightIndent;
  target.lineSpacing = source.lineSpacing;
  target.spaceBefore = source.spaceBefore;
  target.spaceAfter = source.spaceAfter;
  target.font.set({
    name: source.font.name,
    size: source.font.size,
    bold: source.font.bold,
    italic: source.font.italic,
    color: source.font.color,
  });
};

interface RepeatNotesHeaderPassResult {
  notesPagesFound: number;
  inserted: boolean;
  duplicatesSkipped: number;
}

async function repeatNotesHeaderPass(): Promise<RepeatNotesHeaderPassResult> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const pages = context.document.activeWindow.activePane.pages;

    paragraphs.load(
      "items/text,items/style,items/alignment,items/firstLineIndent,items/leftIndent," +
        "items/rightIndent,items/lineSpacing,items/spaceBefore,items/spaceAfter"
    );
    pages.load("items/index");
    await context.sync();

    const paragraphPages = paragraphs.items.map((paragraph) => {
      const collection = paragraph.getRange(Word.RangeLocation.content).pages;
      collection.load("items/index");
      paragraph.font.load("name,size,bold,italic,color");
      paragraph.listItemOrNullObject.load("isNullObject,listString");
      return collection;
    });
    const paragraphStartPages = paragraphs.items.map((paragraph) => {
      const collection = paragraph.getRange(Word.RangeLocation.start).pages;
      collection.load("items/index");
      return collection;
    });
    await context.sync();

    const snapshot: NotesPageParagraph[] = paragraphs.items.map((paragraph, index) => {
      const listItem = paragraph.listItemOrNullObject;
      return {
        text: paragraph.text,
        listPrefix: listItem.isNullObject ? "" : listItem.listString,
        pages: paragraphPages[index].items.map((page) => page.index),
        startPage: paragraphStartPages[index].items[0]?.index ?? null,
      };
    });
    const location = locateNotesSection(snapshot);
    if (!location) {
      throw new Error(
        "Could not find the Notes title, its financial year-end line, and a numbered Notes heading."
      );
    }

    const duplicatePages = location.pages.filter((pageIndex) =>
      pageHasNotesHeader(snapshot, pageIndex)
    );
    const targetPageIndex = location.pages.find(
      (pageIndex) => !pageHasNotesHeader(snapshot, pageIndex)
    );
    if (targetPageIndex === undefined) {
      return {
        notesPagesFound: location.pages.length,
        inserted: false,
        duplicatesSkipped: duplicatePages.length,
      };
    }

    const page = pages.items.find((item) => item.index === targetPageIndex);
    if (!page) throw new Error(`Could not access rendered Notes page ${targetPageIndex}.`);

    const titleSource = paragraphs.items[location.titleIndex];
    const yearEndSource = paragraphs.items[location.yearEndIndex];
    const firstParagraphIndex = snapshot.findIndex(
      (paragraph) => paragraph.startPage === targetPageIndex
    );
    const firstParagraph = paragraphs.items[firstParagraphIndex];
    const startsWithContinuationHeading =
      firstParagraphIndex >= 0 &&
      CONTINUATION_SUFFIX_PATTERN.test(normalizeWordText(snapshot[firstParagraphIndex].text));
    const title = startsWithContinuationHeading
      ? firstParagraph.insertParagraph(titleSource.text.trim(), Word.InsertLocation.before)
      : page
          .getRange(Word.RangeLocation.start)
          .insertParagraph(titleSource.text.trim(), Word.InsertLocation.before);
    const yearEnd = title.insertParagraph(yearEndSource.text.trim(), Word.InsertLocation.after);
    copyParagraphFormatting(titleSource, title);
    copyParagraphFormatting(yearEndSource, yearEnd);

    const headerRange = title
      .getRange(Word.RangeLocation.whole)
      .expandTo(yearEnd.getRange(Word.RangeLocation.whole));
    const control = headerRange.insertContentControl();
    control.tag = `${NOTES_HEADER_TAG_PREFIX}${targetPageIndex}`;
    control.title = "Repeated Notes header";
    control.appearance = Word.ContentControlAppearance.hidden;
    await context.sync();

    return {
      notesPagesFound: location.pages.length,
      inserted: true,
      duplicatesSkipped: duplicatePages.length,
    };
  });
}

export async function repeatNotesHeader(): Promise<RepeatNotesHeaderResult> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(
      "Repeat Notes Header requires WordApiDesktop 1.2 in a supported Word desktop client."
    );
  }

  let headersInserted = 0;
  let notesPagesFound = 0;
  let duplicatesSkipped = 0;
  const maximumPasses = 250;

  for (let pass = 0; pass < maximumPasses; pass += 1) {
    const result = await repeatNotesHeaderPass();
    notesPagesFound = result.notesPagesFound;
    duplicatesSkipped = result.duplicatesSkipped;
    if (!result.inserted) {
      return { notesPagesFound, headersInserted, duplicatesSkipped };
    }
    headersInserted += 1;
  }

  throw new Error("Repeat Notes Header reached its 250-page safety limit.");
}
