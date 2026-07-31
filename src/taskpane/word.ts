/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import {
  addKeepLinesToAllParagraphs,
  disableMoveToNextPage,
  DocumentKeepLinesResult,
  enableKeepTogether,
  MoveParagraphsResult,
  removeKeepLinesFromAllParagraphs,
  RemoveMoveResult,
} from "./paragraph-format";
import {
  continuationText,
  CONTINUATION_SUFFIX_PATTERN,
  parseNumericHeading,
  startsWithNumericHeading,
} from "./continuation-format";

export { continuationText, parseNumericHeading } from "./continuation-format";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMessage = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    const runButton = document.getElementById("run");

    if (sideloadMessage) {
      sideloadMessage.style.display = "none";
    }

    if (appBody) {
      appBody.style.display = "flex";
    }

    if (runButton) {
      runButton.onclick = runWord;
    }
  }
});

export interface ParagraphPageResult {
  text: string;
  pages: number[];
}

export interface DocumentPaginationResult {
  pageCount: number;
  paragraphs: ParagraphPageResult[];
}

export interface ContinuationInsertionResult {
  continuingSectionsFound: number;
  continuationPagesFound: number;
  headingsInserted: number;
  duplicatesSkipped: number;
  limitationMessage: string;
}

interface HeadingDetails {
  key: string;
  level: number;
  text: string;
  paragraph: Word.Paragraph;
}

interface ParagraphDetails {
  paragraph: Word.Paragraph;
  text: string;
  pages: number[];
  heading: HeadingDetails | null;
  isInsertedContinuation: boolean;
}

interface ContinuationPage {
  pageIndex: number;
  anchor: Word.Paragraph;
  pageStartRange: Word.Range;
  detectedHeadings: HeadingDetails[];
  headings: HeadingDetails[];
  existingTexts: Set<string>;
  startsInsideParagraph: boolean;
}

const CONTINUATION_TAG_PREFIX = "word-continuation-heading:";
const LEGACY_MARKER_SHAPE_PREFIX = "word-continuation-marker";

export const PAGINATION_REQUIREMENT_MESSAGE =
  "Page detection requires WordApiDesktop 1.2, which is available in supported Word desktop clients.";

export async function moveSelectedParagraphsToNextPage(): Promise<MoveParagraphsResult> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) {
    throw new Error(
      "Move Paragraph to Next Page requires WordApi 1.5, which is available in supported Word clients."
    );
  }

  return Word.run(async (context) => {
    // Word returns the complete paragraph collection for both a text selection and
    // a collapsed insertion point.
    const paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load("items");
    await context.sync();

    const paragraphRanges = paragraphs.items.map((paragraph) =>
      paragraph.getRange(Word.RangeLocation.whole)
    );
    const paragraphPages = paragraphRanges.map((range) => {
      const pages = range.pages;
      pages.load("items/index");
      return pages;
    });
    const paragraphOoxml = paragraphRanges.map((range) => range.getOoxml());
    await context.sync();

    const update = enableKeepTogether(
      paragraphRanges.map((_range, index) => ({
        ooxml: paragraphOoxml[index].value,
        pageCount: paragraphPages[index].items.length,
      }))
    );

    for (let index = paragraphRanges.length - 1; index >= 0; index -= 1) {
      paragraphRanges[index].insertOoxml(update.paragraphs[index], Word.InsertLocation.replace);
    }
    await context.sync();
    return {
      paragraphsUpdated: update.paragraphsUpdated,
      splitParagraphsFound: update.splitParagraphsFound,
    };
  });
}

export async function removeMoveFromSelectedParagraphs(): Promise<RemoveMoveResult> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) {
    throw new Error(
      "Remove Move to Next Page requires WordApi 1.5, which is available in supported Word clients."
    );
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load("items");
    await context.sync();

    const paragraphRanges = paragraphs.items.map((paragraph) =>
      paragraph.getRange(Word.RangeLocation.whole)
    );
    const paragraphOoxml = paragraphRanges.map((range) => range.getOoxml());
    await context.sync();

    const update = disableMoveToNextPage(paragraphOoxml.map(({ value }) => value));
    for (let index = paragraphRanges.length - 1; index >= 0; index -= 1) {
      paragraphRanges[index].insertOoxml(update.paragraphs[index], Word.InsertLocation.replace);
    }
    await context.sync();
    return update.result;
  });
}

export async function keepAllParagraphsOnOnePage(): Promise<DocumentKeepLinesResult> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.1")) {
    throw new Error("Keep All Paragraphs on One Page requires WordApi 1.1.");
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const bodyOoxml = body.getOoxml();
    await context.sync();

    // ClientResult values become available after context.sync(); they aren't loadable ClientObjects.
    // eslint-disable-next-line office-addins/load-object-before-read
    const update = addKeepLinesToAllParagraphs(bodyOoxml.value);
    if (update.result.paragraphsChanged > 0) {
      body.insertOoxml(update.ooxml, Word.InsertLocation.replace);
      await context.sync();
    }
    return update.result;
  });
}

export async function removeKeepAllParagraphsTogether(): Promise<DocumentKeepLinesResult> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.1")) {
    throw new Error("Remove Keep All Paragraphs Together requires WordApi 1.1.");
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const bodyOoxml = body.getOoxml();
    await context.sync();

    // ClientResult values become available after context.sync(); they aren't loadable ClientObjects.
    // eslint-disable-next-line office-addins/load-object-before-read
    const update = removeKeepLinesFromAllParagraphs(bodyOoxml.value);
    if (update.result.paragraphsChanged > 0) {
      body.insertOoxml(update.ooxml, Word.InsertLocation.replace);
      await context.sync();
    }
    return update.result;
  });
}

const normalizeHeadingText = (text: string): string =>
  text.replace(/\s+/g, " ").trim().toLocaleLowerCase();

const updateHierarchy = (
  hierarchy: Array<HeadingDetails | undefined>,
  heading: HeadingDetails
): void => {
  hierarchy.length = heading.level;
  hierarchy[heading.level - 1] = heading;
};

export async function analyzeDocumentPagination(): Promise<DocumentPaginationResult> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(PAGINATION_REQUIREMENT_MESSAGE);
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const pages = context.document.activeWindow.activePane.pages;

    paragraphs.load("items/text");
    pages.load("items/index");
    await context.sync();

    const paragraphPageCollections = paragraphs.items.map((paragraph) => {
      const paragraphPages = paragraph.getRange().pages;
      paragraphPages.load("items/index");
      return paragraphPages;
    });

    await context.sync();

    return {
      pageCount: pages.items.length,
      paragraphs: paragraphs.items
        .map((paragraph, index) => ({
          text: paragraph.text,
          pages: paragraphPageCollections[index].items.map((page) => page.index),
        }))
        .filter((paragraph) => paragraph.text.trim() !== ""),
    };
  });
}

export async function assessContinuationMarkers(
  insertContinuationHeadings: boolean
): Promise<ContinuationInsertionResult> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(PAGINATION_REQUIREMENT_MESSAGE);
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const pages = context.document.activeWindow.activePane.pages;
    const contentControls = context.document.body.contentControls;

    paragraphs.load(
      "items/text,items/style,items/styleBuiltIn,items/outlineLevel,items/isListItem," +
        "items/alignment,items/firstLineIndent,items/leftIndent,items/rightIndent," +
        "items/lineSpacing,items/spaceBefore,items/spaceAfter"
    );
    pages.load("items/index");
    contentControls.load("items/tag");
    await context.sync();

    const paragraphPageCollections = paragraphs.items.map((paragraph) => {
      const paragraphPages = paragraph.getRange().pages;
      paragraphPages.load("items/index");
      paragraph.font.load("name,size,bold,italic,color");
      paragraph.listItemOrNullObject.load("isNullObject,listString,level");
      return paragraphPages;
    });
    await context.sync();

    const insertedContinuationTags = new Set(
      contentControls.items
        .map((control) => control.tag)
        .filter((tag) => tag.startsWith(CONTINUATION_TAG_PREFIX))
    );

    const paragraphDetails: ParagraphDetails[] = paragraphs.items
      .map((paragraph, index) => {
        const text = paragraph.text.trim();
        const listItem = paragraph.listItemOrNullObject;
        const listPrefix = !listItem.isNullObject ? listItem.listString.trim() : "";
        const numeric = parseNumericHeading(text) ?? parseNumericHeading(`${listPrefix} `);
        const fullHeadingText =
          numeric && listPrefix && !startsWithNumericHeading(text) ? `${listPrefix} ${text}` : text;
        return {
          paragraph,
          text,
          pages: paragraphPageCollections[index].items.map((page) => page.index),
          heading: numeric
            ? { key: numeric.key, level: numeric.level, text: fullHeadingText, paragraph }
            : null,
          isInsertedContinuation: CONTINUATION_SUFFIX_PATTERN.test(text),
        };
      })
      .filter((details) => details.text !== "");

    const originalParagraphs = paragraphDetails.filter(
      (details) => !details.isInsertedContinuation
    );
    const activeHierarchy: Array<HeadingDetails | undefined> = [];
    const hierarchyBeforeParagraph = new Map<Word.Paragraph, HeadingDetails[]>();

    for (const details of originalParagraphs) {
      hierarchyBeforeParagraph.set(
        details.paragraph,
        activeHierarchy.filter((heading): heading is HeadingDetails => Boolean(heading))
      );
      if (details.heading) updateHierarchy(activeHierarchy, details.heading);
    }

    const continuationPages: ContinuationPage[] = [];
    for (const page of pages.items.slice(1)) {
      const paragraphsOnPage = originalParagraphs.filter((details) =>
        details.pages.includes(page.index)
      );
      const firstOnPage = paragraphsOnPage[0];
      if (!firstOnPage) continue;

      const startedEarlier = firstOnPage.pages.some((pageIndex) => pageIndex < page.index);
      if (!startedEarlier && firstOnPage.heading) continue;

      const hierarchy = [...(hierarchyBeforeParagraph.get(firstOnPage.paragraph) ?? [])];
      if (firstOnPage.heading && startedEarlier) updateHierarchy(hierarchy, firstOnPage.heading);
      if (hierarchy.length === 0) continue;

      continuationPages.push({
        pageIndex: page.index,
        anchor: firstOnPage.paragraph,
        pageStartRange: page.getRange(Word.RangeLocation.start),
        detectedHeadings: hierarchy,
        headings: hierarchy.slice(-1),
        existingTexts: new Set(
          paragraphDetails
            .filter((details) => details.pages.includes(page.index))
            .map((details) => normalizeHeadingText(details.text))
        ),
        startsInsideParagraph: startedEarlier,
      });
    }

    const continuingSectionKeys = new Set<string>();
    let headingsInserted = 0;
    let duplicatesSkipped = 0;

    // Work from the last rendered page toward the first so inserting at a page
    // boundary doesn't invalidate the ranges for later continuation pages.
    for (const continuationPage of [...continuationPages].reverse()) {
      for (const heading of continuationPage.detectedHeadings) {
        continuingSectionKeys.add(heading.key);
      }
      for (const heading of continuationPage.headings) {
        const text = continuationText(heading.text);
        const normalizedText = normalizeHeadingText(text);
        const tag = `${CONTINUATION_TAG_PREFIX}${continuationPage.pageIndex}:${heading.key}`;

        if (
          continuationPage.existingTexts.has(normalizedText) ||
          insertedContinuationTags.has(tag)
        ) {
          duplicatesSkipped += 1;
          continue;
        }
        if (!insertContinuationHeadings) continue;

        const inserted = continuationPage.startsInsideParagraph
          ? continuationPage.pageStartRange.insertParagraph(text, Word.InsertLocation.before)
          : continuationPage.anchor.insertParagraph(text, Word.InsertLocation.before);
        inserted.style = heading.paragraph.style;
        inserted.alignment = heading.paragraph.alignment;
        inserted.firstLineIndent = heading.paragraph.firstLineIndent;
        inserted.leftIndent = heading.paragraph.leftIndent;
        inserted.rightIndent = heading.paragraph.rightIndent;
        inserted.lineSpacing = heading.paragraph.lineSpacing;
        inserted.spaceBefore = heading.paragraph.spaceBefore;
        inserted.spaceAfter = heading.paragraph.spaceAfter;
        inserted.font.set({
          name: heading.paragraph.font.name,
          size: heading.paragraph.font.size,
          bold: heading.paragraph.font.bold,
          italic: heading.paragraph.font.italic,
          color: heading.paragraph.font.color,
        });

        const control = inserted.insertContentControl();
        control.tag = tag;
        control.title = "Continuation heading";
        control.appearance = Word.ContentControlAppearance.hidden;
        insertedContinuationTags.add(tag);
        continuationPage.existingTexts.add(normalizedText);
        headingsInserted += 1;
      }
    }

    if (headingsInserted > 0) {
      await context.sync();
      const repaginatedPages = context.document.activeWindow.activePane.pages;
      repaginatedPages.load("items/index");
      await context.sync();
    }

    return {
      continuingSectionsFound: continuingSectionKeys.size,
      continuationPagesFound: continuationPages.length,
      headingsInserted,
      duplicatesSkipped,
      limitationMessage:
        "Word repaginates dynamically after insertion. One normal continuation heading is inserted at the start of each continuation page; if the page starts inside a split paragraph, Word creates a paragraph boundary there.",
    };
  });
}

export async function removeContinuationMarkers(): Promise<number> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(PAGINATION_REQUIREMENT_MESSAGE);
  }

  return Word.run(async (context) => {
    const contentControls = context.document.body.contentControls;
    const shapes = context.document.body.shapes;
    contentControls.load("items/tag");
    shapes.load("items/name");
    await context.sync();

    const continuationControls = contentControls.items.filter((control) =>
      control.tag.startsWith(CONTINUATION_TAG_PREFIX)
    );
    const legacyShapes = shapes.items.filter((shape) =>
      shape.name.startsWith(`${LEGACY_MARKER_SHAPE_PREFIX}-`)
    );

    for (const control of continuationControls) control.delete(false);
    for (const shape of legacyShapes) shape.delete();

    await context.sync();
    return continuationControls.length + legacyShapes.length;
  });
}

export async function runWord() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
