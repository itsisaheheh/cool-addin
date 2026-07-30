/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

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
  boundariesFound: number;
  markersInserted: number;
  limitationMessage: string;
}

const CONTINUED_FROM_MARKER = "(CONT')";
const CONTINUED_ON_MARKER = "(CONT'D)";
const MARKER_SHAPE_PREFIX = "word-continuation-marker";
const MARKER_WIDTH = 96;
const MARKER_HEIGHT = 18;
const PAGE_EDGE_OFFSET = 18;

export const PAGINATION_REQUIREMENT_MESSAGE =
  "Page detection requires WordApiDesktop 1.2, which is available in supported Word desktop clients.";

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

    const paragraphResults = paragraphs.items
      .map((paragraph, index) => ({
        text: paragraph.text,
        pages: paragraphPageCollections[index].items.map((page) => page.index),
      }))
      .filter((paragraph) => paragraph.text.trim() !== "");

    return {
      pageCount: pages.items.length,
      paragraphs: paragraphResults,
    };
  });
}

export async function assessContinuationMarkers(
  includeContinuedOnMarker: boolean
): Promise<ContinuationInsertionResult> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(PAGINATION_REQUIREMENT_MESSAGE);
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const pages = context.document.activeWindow.activePane.pages;
    const shapes = context.document.body.shapes;

    paragraphs.load("items/text");
    pages.load("items/index,items/height,items/width");
    shapes.load("items/name");
    await context.sync();

    const paragraphPageCollections = paragraphs.items.map((paragraph) => {
      const paragraphPages = paragraph.getRange().pages;
      paragraphPages.load("items/index");
      return paragraphPages;
    });

    await context.sync();

    const continuedParagraphs = paragraphs.items
      .map((paragraph, index) => ({
        paragraph,
        pageIndexes: paragraphPageCollections[index].items.map((page) => page.index),
      }))
      .filter(
        ({ paragraph, pageIndexes }) => paragraph.text.trim() !== "" && pageIndexes.length > 1
      );

    const boundariesFound = continuedParagraphs.reduce(
      (total, { pageIndexes }) => total + pageIndexes.length - 1,
      0
    );
    const existingShapeNames = new Set(shapes.items.map((shape) => shape.name));
    const pagesByIndex = new Map(pages.items.map((page) => [page.index, page]));
    let markersInserted = 0;

    const insertMarker = (
      pageIndex: number,
      marker: string,
      markerKind: "from" | "on",
      atBottom: boolean
    ): void => {
      const shapeName = `${MARKER_SHAPE_PREFIX}-${markerKind}-page-${pageIndex}`;
      if (existingShapeNames.has(shapeName)) {
        return;
      }

      const page = pagesByIndex.get(pageIndex);
      if (!page) {
        return;
      }

      const shape = page.getRange(Word.RangeLocation.start).insertTextBox(marker, {
        width: MARKER_WIDTH,
        height: MARKER_HEIGHT,
      });

      shape.name = shapeName;
      shape.altTextDescription = `Continuation marker ${marker} for page ${pageIndex}`;
      shape.relativeHorizontalPosition = Word.RelativeHorizontalPosition.page;
      shape.relativeVerticalPosition = Word.RelativeVerticalPosition.page;
      shape.left = Math.max(0, (page.width - MARKER_WIDTH) / 2);
      shape.top = atBottom
        ? Math.max(0, page.height - MARKER_HEIGHT - PAGE_EDGE_OFFSET)
        : PAGE_EDGE_OFFSET;
      shape.allowOverlap = true;
      shape.fill.clear();
      shape.textWrap.type = Word.ShapeTextWrapType.front;
      shape.textFrame.set({
        autoSizeSetting: Word.ShapeAutoSize.none,
        bottomMargin: 0,
        leftMargin: 0,
        rightMargin: 0,
        topMargin: 0,
        verticalAlignment: Word.ShapeTextVerticalAlignment.middle,
        wordWrap: false,
      });
      shape.body.font.set({
        bold: true,
        name: "Arial",
        size: 10,
      });
      shape.body.paragraphs.getFirst().alignment = Word.Alignment.centered;

      existingShapeNames.add(shapeName);
      markersInserted += 1;
    };

    for (const { pageIndexes } of continuedParagraphs) {
      for (let index = 0; index < pageIndexes.length - 1; index += 1) {
        insertMarker(pageIndexes[index], CONTINUED_FROM_MARKER, "from", true);

        if (includeContinuedOnMarker) {
          insertMarker(pageIndexes[index + 1], CONTINUED_ON_MARKER, "on", false);
        }
      }
    }

    await context.sync();

    return {
      boundariesFound,
      markersInserted,
      limitationMessage:
        "Markers are floating text boxes anchored to their pages, so the original paragraph text and pagination are unchanged.",
    };
  });
}

export async function removeContinuationMarkers(): Promise<number> {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2")) {
    throw new Error(PAGINATION_REQUIREMENT_MESSAGE);
  }

  return Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items/name");
    await context.sync();

    const markerShapes = shapes.items.filter((shape) =>
      shape.name.startsWith(`${MARKER_SHAPE_PREFIX}-`)
    );

    for (const shape of markerShapes) {
      shape.delete();
    }

    await context.sync();
    return markerShapes.length;
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
