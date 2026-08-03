/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office, Word */

import {
  addKeepNextToParagraphOoxml,
  addKeepLinesToAllParagraphs,
  disableMoveToNextPage,
  DocumentKeepLinesResult,
  enableKeepTogether,
  MoveParagraphsResult,
  removeKeepLinesFromAllParagraphs,
  RemoveMoveResult,
  validateKeepLinesPagination,
} from "./paragraph-format";
import {
  continuationText,
  continuationPageEligibility,
  continuationPlacement,
  CONTINUATION_SUFFIX_PATTERN,
  isOrphanOriginalHeading,
  parseNumericHeading,
  startsWithNumericHeading,
} from "./continuation-format";
import { runContdInsertionUntilStable } from "./contd-stabilization";

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

export interface KeepParagraphsIntactResult extends DocumentKeepLinesResult {
  splitParagraphsFixed: number;
  orphanHeadingsFixed: number;
  paginationPasses: number;
  unfixableParagraphs: number;
}

export interface ContinuationInsertionResult {
  continuingSectionsFound: number;
  continuationPagesFound: number;
  headingsInserted: number;
  duplicatesSkipped: number;
  limitationMessage: string;
  paginationChanged?: boolean;
}

export interface DocumentCheckResult {
  numberedHeadings: NumberedHeadingCheckResult[];
}

export interface NumberedHeadingCheckResult {
  key: string;
  title: string;
}

interface HeadingDetails {
  key: string;
  level: number;
  text: string;
  paragraph: Word.Paragraph;
  startPage: number | null;
}

interface ParagraphDetails {
  documentIndex: number;
  paragraph: Word.Paragraph;
  text: string;
  pages: number[];
  startPage: number | null;
  heading: HeadingDetails | null;
  isInsertedContinuation: boolean;
}

interface ContinuationPage {
  pageIndex: number;
  pageStartRange: Word.Range;
  anchorIndex: number;
  anchorText: string;
  anchor: Word.Paragraph;
  detectedHeadings: HeadingDetails[];
  headings: HeadingDetails[];
  existingTexts: Set<string>;
  startsInsideParagraph: boolean;
}

const CONTINUATION_TAG_PREFIX = "word-continuation-heading:";
const LEGACY_MARKER_SHAPE_PREFIX = "word-continuation-marker";

const logContdDiagnostic = (details: Record<string, unknown>): void => {
  console.debug("[Add CONT'D Headings]", details);
};

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

export async function keepAllParagraphsOnOnePage(): Promise<KeepParagraphsIntactResult> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(PAGINATION_REQUIREMENT_MESSAGE);
  }

  return Word.run(async (context) => {
    const initialParagraphs = context.document.body.paragraphs;
    initialParagraphs.load("items");
    await context.sync();

    const totalParagraphCount = initialParagraphs.items.length;
    const validation = await validateKeepLinesPagination(totalParagraphCount, async () => {
      // Every pass creates new ranges and reloads page membership after the
      // preceding context.sync() has allowed Word to repaginate.
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const ranges = paragraphs.items.map((paragraph) =>
        paragraph.getRange(Word.RangeLocation.whole)
      );
      const pageCollections = ranges.map((range) => {
        const pages = range.pages;
        pages.load("items/index");
        return pages;
      });
      const paragraphOoxml = ranges.map((range) => range.getOoxml());
      await context.sync();

      return {
        paragraphs: paragraphs.items.map((paragraph, index) => ({
          text: paragraph.text,
          ooxml: paragraphOoxml[index].value,
          pageCount: pageCollections[index].items.length,
          pages: pageCollections[index].items.map((page) => page.index),
        })),
        applyParagraphs: async (updates) => {
          for (const update of [...updates].sort((left, right) => right.index - left.index)) {
            ranges[update.index].insertOoxml(update.ooxml, Word.InsertLocation.replace);
          }
          await context.sync();
        },
        settlePagination: async () => {
          const renderedPages = context.document.activeWindow.activePane.pages;
          renderedPages.load("items/index");
          await context.sync();
        },
      };
    });

    return {
      paragraphsFound: validation.paragraphsChecked,
      paragraphsChanged: validation.splitParagraphsFixed,
      paragraphsAlreadyFormatted: validation.unfixableParagraphs,
      splitParagraphsFixed: validation.splitParagraphsFixed,
      orphanHeadingsFixed: validation.orphanHeadingsFixed,
      paginationPasses: validation.paginationPasses,
      unfixableParagraphs: validation.unfixableParagraphs,
    };
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

export async function checkDocumentIssues(): Promise<DocumentCheckResult> {
  const numberedHeadings = await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    paragraphs.items.forEach((paragraph) => {
      paragraph.listItemOrNullObject.load("isNullObject,listString");
    });
    await context.sync();

    return paragraphs.items.flatMap((paragraph) => {
      const listItem = paragraph.listItemOrNullObject;
      const listPrefix = listItem.isNullObject ? "" : listItem.listString.trim();
      const numeric = parseNumericHeading(paragraph.text) ?? parseNumericHeading(`${listPrefix} `);
      if (!numeric) return [];

      const text =
        listPrefix && !startsWithNumericHeading(paragraph.text)
          ? `${listPrefix} ${paragraph.text.trim()}`
          : paragraph.text.trim();
      return [
        {
          key: numeric.key,
          title: text.replace(/^\d+(?:\.\d+)*\.?\s*/, "").trim() || "Untitled section",
        },
      ];
    });
  });
  return { numberedHeadings };
}

export async function addContdHeadings(): Promise<ContinuationInsertionResult> {
  const paragraphCount = await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    return paragraphs.items.length;
  });

  return runContdInsertionUntilStable(Math.max(2, paragraphCount + 2), async (passNumber) =>
    assessContinuationMarkers(true, passNumber)
  );
}

interface ContinuationPassResult extends ContinuationInsertionResult {
  requiresRepaginationPass: boolean;
}

export async function assessContinuationMarkers(
  insertContinuationHeadings: boolean,
  diagnosticPass = 1
): Promise<ContinuationInsertionResult> {
  const preparation = await assessContinuationMarkersPass(
    insertContinuationHeadings,
    true,
    diagnosticPass
  );
  if (!preparation.requiresRepaginationPass) return preparation;

  const repaginated = await assessContinuationMarkersPass(
    insertContinuationHeadings,
    false,
    diagnosticPass
  );
  return {
    ...repaginated,
    paginationChanged: preparation.paginationChanged || repaginated.paginationChanged,
  };
}

async function assessContinuationMarkersPass(
  insertContinuationHeadings: boolean,
  prepareAffectedParagraphs: boolean,
  diagnosticPass: number
): Promise<ContinuationPassResult> {
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
    const pageCountBefore = pages.items.length;

    const paragraphPageCollections = paragraphs.items.map((paragraph) => {
      // Exclude the paragraph mark: Word can render that mark on the following
      // page even when all visible paragraph content remains on the current page.
      const paragraphPages = paragraph.getRange(Word.RangeLocation.content).pages;
      paragraphPages.load("items/index");
      paragraph.font.load("name,size,bold,italic,color");
      paragraph.listItemOrNullObject.load("isNullObject,listString,level");
      return paragraphPages;
    });
    const paragraphStartPageCollections = paragraphs.items.map((paragraph) => {
      const startPages = paragraph.getRange(Word.RangeLocation.start).pages;
      startPages.load("items/index");
      return startPages;
    });
    await context.sync();

    const insertedContinuationTags = new Set(
      contentControls.items
        .map((control) => control.tag)
        .filter(
          (tag): tag is string => typeof tag === "string" && tag.startsWith(CONTINUATION_TAG_PREFIX)
        )
    );

    const paragraphDetails: ParagraphDetails[] = paragraphs.items
      .map((paragraph, index) => {
        const text = paragraph.text.trim();
        const listItem = paragraph.listItemOrNullObject;
        const listPrefix = !listItem.isNullObject ? listItem.listString.trim() : "";
        const numeric = parseNumericHeading(text) ?? parseNumericHeading(`${listPrefix} `);
        const fullHeadingText =
          numeric && listPrefix && !startsWithNumericHeading(text) ? `${listPrefix} ${text}` : text;
        const startPage = paragraphStartPageCollections[index].items[0]?.index ?? null;
        return {
          documentIndex: index,
          paragraph,
          text,
          pages: paragraphPageCollections[index].items.map((page) => page.index),
          startPage,
          heading: numeric
            ? {
                key: numeric.key,
                level: numeric.level,
                text: fullHeadingText,
                paragraph,
                startPage,
              }
            : null,
          isInsertedContinuation: CONTINUATION_SUFFIX_PATTERN.test(text),
        };
      })
      .filter((details) => details.text !== "");

    const originalParagraphs = paragraphDetails.filter(
      (details) => !details.isInsertedContinuation
    );
    const orphanHeadingPairs = originalParagraphs.flatMap((details, index) => {
      const nextParagraph = originalParagraphs[index + 1];
      if (
        !details.heading ||
        !nextParagraph ||
        !isOrphanOriginalHeading({
          headingStartPage: details.startPage,
          nextContentStartPage: nextParagraph.startPage,
          nextParagraphIsNumberedHeading: nextParagraph.heading !== null,
        })
      ) {
        return [];
      }
      return [{ heading: details, content: nextParagraph }];
    });

    if (prepareAffectedParagraphs && insertContinuationHeadings && orphanHeadingPairs.length > 0) {
      const firstOrphan = orphanHeadingPairs[0];
      const headingRange = firstOrphan.heading.paragraph.getRange(Word.RangeLocation.whole);
      const headingOoxml = headingRange.getOoxml();
      await context.sync();

      // ClientResult values become available after context.sync(); they aren't loadable ClientObjects.
      // eslint-disable-next-line office-addins/load-object-before-read
      const originalHeadingOoxml = headingOoxml.value;
      const updatedHeadingOoxml = addKeepNextToParagraphOoxml(originalHeadingOoxml);
      const headingChanged = updatedHeadingOoxml !== originalHeadingOoxml;
      if (headingChanged) {
        headingRange.insertOoxml(updatedHeadingOoxml, Word.InsertLocation.replace);
        await context.sync();
      }

      logContdDiagnostic({
        pass: diagnosticPass,
        phase: "orphan-heading",
        totalParagraphs: paragraphs.items.length,
        candidateSection: firstOrphan.heading.heading?.key ?? null,
        candidatePage: firstOrphan.heading.startPage,
        candidateParagraphIndex: firstOrphan.heading.documentIndex,
        insertionTargetText: firstOrphan.content.text,
        paginationChanged: headingChanged,
        reason: headingChanged
          ? "Applied keepNext to the orphan original heading; rescanning before any CONT'D insertion."
          : "Orphan original heading already has keepNext; suppressing CONT'D on its first content page.",
      });

      if (headingChanged) {
        return {
          continuingSectionsFound: 0,
          continuationPagesFound: 0,
          headingsInserted: 0,
          duplicatesSkipped: 0,
          limitationMessage:
            "Moved an orphan original heading with its first content paragraph and repaginated.",
          paginationChanged: true,
          requiresRepaginationPass: true,
        };
      }
    }

    const unresolvedOrphanContent = new Set(
      orphanHeadingPairs.map(({ content }) => content.paragraph)
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
      const paragraphSpanningIntoPage = paragraphsOnPage.find(
        (details) =>
          details.startPage !== null &&
          details.startPage < page.index &&
          details.pages.includes(page.index)
      );
      const firstStartingOnPage = originalParagraphs.find(
        (details) => details.startPage === page.index
      );
      const firstOnPage = paragraphSpanningIntoPage ?? firstStartingOnPage;
      if (!firstOnPage) continue;
      if (unresolvedOrphanContent.has(firstOnPage.paragraph)) {
        logContdDiagnostic({
          pass: diagnosticPass,
          phase: "candidate",
          totalParagraphs: paragraphs.items.length,
          candidateSection:
            hierarchyBeforeParagraph.get(firstOnPage.paragraph)?.slice(-1)[0]?.key ?? null,
          candidatePage: page.index,
          candidateParagraphIndex: firstOnPage.documentIndex,
          insertionTargetText: firstOnPage.text,
          paginationChanged: false,
          reason: "Skipped CONT'D because this is the first content of an orphan original heading.",
        });
        continue;
      }

      const startedEarlier = paragraphSpanningIntoPage === firstOnPage;
      const hierarchy = [...(hierarchyBeforeParagraph.get(firstOnPage.paragraph) ?? [])].filter(
        (heading) => heading.startPage !== null && heading.startPage < page.index
      );
      if (hierarchy.length === 0) continue;
      const activeHeading = hierarchy[hierarchy.length - 1];
      const eligibility = continuationPageEligibility({
        sectionStartPage: activeHeading.startPage ?? page.index,
        currentPage: page.index,
        anchorStartPage: firstOnPage.startPage,
        anchorIsOriginalHeading: firstOnPage.heading !== null,
        anchorSpansFromEarlierPage: startedEarlier,
      });
      if (eligibility === "skip") continue;

      continuationPages.push({
        pageIndex: page.index,
        pageStartRange: page.getRange(Word.RangeLocation.start),
        anchorIndex: firstOnPage.documentIndex,
        anchorText: firstOnPage.text,
        anchor: firstOnPage.paragraph,
        detectedHeadings: hierarchy,
        headings: hierarchy.slice(-1),
        existingTexts: new Set(
          paragraphDetails
            .filter((details) => details.pages.includes(page.index))
            .map((details) => normalizeHeadingText(details.text))
        ),
        startsInsideParagraph: eligibility === "prepare",
      });
    }

    const continuingSectionKeys = new Set<string>();
    let headingsInserted = 0;
    let duplicatesSkipped = 0;

    if (prepareAffectedParagraphs && insertContinuationHeadings) {
      const firstCandidate = continuationPages[0];
      if (firstCandidate) {
        const affectedRange = firstCandidate.anchor.getRange(Word.RangeLocation.whole);
        const affectedOoxml = affectedRange.getOoxml();
        await context.sync();

        // ClientResult values become available after context.sync(); they aren't loadable ClientObjects.
        // eslint-disable-next-line office-addins/load-object-before-read
        const update = addKeepLinesToAllParagraphs(affectedOoxml.value);
        if (update.result.paragraphsChanged > 0) {
          affectedRange.insertOoxml(update.ooxml, Word.InsertLocation.replace);
          await context.sync();
        }

        const candidateHeading = firstCandidate.headings[0];
        logContdDiagnostic({
          pass: diagnosticPass,
          phase: "prepare",
          totalParagraphs: paragraphs.items.length,
          candidateSection: candidateHeading?.key ?? null,
          candidatePage: firstCandidate.pageIndex,
          candidateParagraphIndex: firstCandidate.anchorIndex,
          insertionTargetText: firstCandidate.anchorText,
          paginationChanged: update.result.paragraphsChanged > 0,
          reason:
            update.result.paragraphsChanged > 0
              ? "Prepared the first candidate paragraph and discarded this pass's layout."
              : "First candidate was already prepared; rescanning before insertion.",
        });
        return {
          continuingSectionsFound: 0,
          continuationPagesFound: continuationPages.length,
          headingsInserted: 0,
          duplicatesSkipped: 0,
          limitationMessage: "Affected paragraphs were prepared for repagination.",
          paginationChanged: update.result.paragraphsChanged > 0,
          requiresRepaginationPass: true,
        };
      }
    }

    const insertedHeadingParagraphs: Array<{
      paragraph: Word.Paragraph;
      tag: string;
      continuationPage: ContinuationPage;
    }> = [];
    let paragraphsStillSplitAfterValidation = 0;

    // Insert at most one heading. Word repaginates after the sync below, and the
    // stabilization loop then discards every candidate and rescans from scratch.
    for (const continuationPage of continuationPages) {
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
          logContdDiagnostic({
            pass: diagnosticPass,
            phase: "insert",
            totalParagraphs: paragraphs.items.length,
            candidateSection: heading.key,
            candidatePage: continuationPage.pageIndex,
            candidateParagraphIndex: continuationPage.anchorIndex,
            insertionTargetText: continuationPage.anchorText,
            paginationChanged: false,
            reason: "Skipped because a matching CONT'D heading already exists.",
          });
          continue;
        }
        if (!insertContinuationHeadings) continue;

        const placement = continuationPlacement(continuationPage.startsInsideParagraph, false);
        const inserted =
          placement === "at-rendered-page-start"
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

        insertedHeadingParagraphs.push({ paragraph: inserted, tag, continuationPage });
        insertedContinuationTags.add(tag);
        continuationPage.existingTexts.add(normalizedText);
        headingsInserted += 1;
        logContdDiagnostic({
          pass: diagnosticPass,
          phase: "insert-target",
          totalParagraphs: paragraphs.items.length,
          candidateSection: heading.key,
          candidatePage: continuationPage.pageIndex,
          candidateParagraphIndex: continuationPage.anchorIndex,
          insertionTargetText: continuationPage.anchorText,
          paginationChanged: false,
          reason:
            placement === "at-rendered-page-start"
              ? "Paragraph still spans pages; inserting at the freshly loaded rendered page start."
              : "Inserting before the first complete paragraph on the continuation page.",
        });
        break;
      }
      if (headingsInserted > 0) break;
    }

    if (insertedHeadingParagraphs.length > 0) {
      await context.sync();

      const insertedRanges = insertedHeadingParagraphs.map(({ paragraph }) =>
        paragraph.getRange(Word.RangeLocation.whole)
      );
      const insertedOoxml = insertedRanges.map((range) => range.getOoxml());
      await context.sync();

      for (let index = insertedRanges.length - 1; index >= 0; index -= 1) {
        const update = addKeepLinesToAllParagraphs(insertedOoxml[index].value);
        const headingOoxml = addKeepNextToParagraphOoxml(update.ooxml);
        const replacement = insertedRanges[index].insertOoxml(
          headingOoxml,
          Word.InsertLocation.replace
        );
        const control = replacement.insertContentControl();
        control.tag = insertedHeadingParagraphs[index].tag;
        control.title = "Continuation heading";
        control.appearance = Word.ContentControlAppearance.hidden;
      }
      await context.sync();

      // Required post-insertion pagination validation. This is deliberately one
      // bounded validation pass: keepLines moves paragraphs that can fit, while
      // paragraphs longer than a page remain split without causing a loop.
      const affectedPages = insertedHeadingParagraphs.map(({ continuationPage }) => {
        const paragraphPages = continuationPage.anchor.getRange().pages;
        paragraphPages.load("items/index");
        return paragraphPages;
      });
      const repaginatedPages = context.document.activeWindow.activePane.pages;
      repaginatedPages.load("items/index");
      await context.sync();
      const insertedCandidate = insertedHeadingParagraphs[0].continuationPage;
      logContdDiagnostic({
        pass: diagnosticPass,
        phase: "insert",
        totalParagraphs: paragraphs.items.length,
        candidateSection: insertedCandidate.headings[0]?.key ?? null,
        candidatePage: insertedCandidate.pageIndex,
        candidateParagraphIndex: insertedCandidate.anchorIndex,
        insertionTargetText: insertedCandidate.anchorText,
        paginationChanged: repaginatedPages.items.length !== pageCountBefore,
        reason: "Inserted one CONT'D heading; all saved candidates will now be discarded.",
      });
      paragraphsStillSplitAfterValidation = affectedPages.filter(
        (paragraphPages) => paragraphPages.items.length > 1
      ).length;
    }

    return {
      continuingSectionsFound: continuingSectionKeys.size,
      continuationPagesFound: continuationPages.length,
      headingsInserted,
      duplicatesSkipped,
      limitationMessage: `Word repaginated after insertion. Affected body paragraphs use Keep lines together, and continuation headings use Keep with next so a heading is followed by the complete paragraph whenever that paragraph can fit on one page. ${paragraphsStillSplitAfterValidation} affected paragraph(s) remain longer than or unable to fit on one page.`,
      requiresRepaginationPass: false,
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
      typeof control.tag === "string" ? control.tag.startsWith(CONTINUATION_TAG_PREFIX) : false
    );
    const legacyShapes = shapes.items.filter((shape) =>
      typeof shape.name === "string"
        ? shape.name.startsWith(`${LEGACY_MARKER_SHAPE_PREFIX}-`)
        : false
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
