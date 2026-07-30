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
