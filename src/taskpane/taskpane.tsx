import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";
import {
  analyzeDocumentPagination,
  assessContinuationMarkers,
  keepAllParagraphsOnOnePage,
  ParagraphPageResult,
  removeContinuationMarkers,
  removeKeepAllParagraphsTogether,
} from "./word";

const formatPageLabel = (pages: number[]): string => {
  if (pages.length === 0) {
    return "Page unavailable";
  }

  if (pages.length === 1) {
    return `Page ${pages[0]}`;
  }

  return `Pages ${pages.join("–")}`;
};

const App = (): React.ReactElement => {
  const [status, setStatus] = React.useState("Ready");
  const [pageCount, setPageCount] = React.useState<number | null>(null);
  const [paragraphs, setParagraphs] = React.useState<ParagraphPageResult[]>([]);
  const [insertContinuationHeadings, setInsertContinuationHeadings] = React.useState(true);
  const [isChecking, setIsChecking] = React.useState(false);

  const checkDocument = async (): Promise<void> => {
    setStatus("Checking document...");
    setIsChecking(true);
    setPageCount(null);
    setParagraphs([]);

    try {
      const insertionResult = await assessContinuationMarkers(insertContinuationHeadings);
      const paginationResult = await analyzeDocumentPagination();

      setPageCount(paginationResult.pageCount);
      setParagraphs(paginationResult.paragraphs);
      setStatus(
        `Found ${insertionResult.continuingSectionsFound} continuing sections across ${insertionResult.continuationPagesFound} continuation pages. Inserted ${insertionResult.headingsInserted} continuation headings and skipped ${insertionResult.duplicatesSkipped} duplicates. ${insertionResult.limitationMessage}`
      );

      console.log("Continuation insertion:", insertionResult);
      console.log("Document pagination:", paginationResult);
    } catch (error) {
      console.error("Word error:", error);

      const message = error instanceof Error ? error.message : JSON.stringify(error);

      setStatus(`Error: ${message}`);
    } finally {
      setIsChecking(false);
    }
  };

  const undoMarkers = async (): Promise<void> => {
    setStatus("Removing continuation markers...");
    setIsChecking(true);

    try {
      const removedCount = await removeContinuationMarkers();
      const paginationResult = await analyzeDocumentPagination();

      setPageCount(paginationResult.pageCount);
      setParagraphs(paginationResult.paragraphs);
      setStatus(`Removed ${removedCount} continuation headings or legacy markers.`);
    } catch (error) {
      console.error("Word error:", error);

      const message = error instanceof Error ? error.message : JSON.stringify(error);
      setStatus(`Error: ${message}`);
    } finally {
      setIsChecking(false);
    }
  };

  const keepAllParagraphsTogether = async (): Promise<void> => {
    setStatus("Applying paragraph pagination formatting...");
    setIsChecking(true);

    try {
      const result = await keepAllParagraphsOnOnePage();
      setStatus(
        `Keep Paragraphs Intact completed. Applied Keep lines together to ${result.splitParagraphsFixed} split body paragraphs and Keep with next to ${result.headingsKept} immediate lettered topic headings.`
      );
    } catch (error) {
      console.error("Word error:", error);
      const message = error instanceof Error ? error.message : JSON.stringify(error);
      setStatus(`Error applying paragraph pagination formatting: ${message}`);
    } finally {
      setIsChecking(false);
    }
  };

  const removeKeepAllParagraphs = async (): Promise<void> => {
    setStatus("Removing paragraph pagination formatting...");
    setIsChecking(true);

    try {
      const result = await removeKeepAllParagraphsTogether();
      setStatus(
        `Success: removed Keep lines together from ${result.paragraphsChanged} of ${result.paragraphsFound} paragraphs. Other formatting was preserved.`
      );
    } catch (error) {
      console.error("Word error:", error);
      const message = error instanceof Error ? error.message : JSON.stringify(error);
      setStatus(`Error removing paragraph pagination formatting: ${message}`);
    } finally {
      setIsChecking(false);
    }
  };

  return (
    <main className="app">
      <h1>Document Continuation Checker</h1>

      <p>Check whether document content continues onto another page.</p>

      <label className="marker-option">
        <input
          type="checkbox"
          checked={insertContinuationHeadings}
          onChange={(event) => setInsertContinuationHeadings(event.target.checked)}
          disabled={isChecking}
        />
        Insert repeated numeric headings with (Cont&apos;d)
      </label>

      <div className="button-row">
        <button type="button" onClick={checkDocument} disabled={isChecking}>
          {isChecking ? "Working..." : "Check Document"}
        </button>
        <button
          className="undo-button"
          type="button"
          onClick={undoMarkers}
          disabled={isChecking}
          aria-label="Undo continuation markers"
          title="Undo continuation markers"
        >
          ↶
        </button>
      </div>

      <div className="button-row">
        <button type="button" onClick={keepAllParagraphsTogether} disabled={isChecking}>
          Keep Paragraphs Intact
        </button>
        <button
          className="undo-button"
          type="button"
          onClick={removeKeepAllParagraphs}
          disabled={isChecking}
          aria-label="Undo Keep Paragraphs Intact"
          title="Undo Keep Paragraphs Intact"
        >
          {"\u21B6"}
        </button>
      </div>

      <p>
        <strong>Status:</strong> {status}
      </p>

      {pageCount !== null && (
        <p className="page-summary">
          <strong>Total pages:</strong> {pageCount}
        </p>
      )}

      {paragraphs.length > 0 && (
        <section>
          <h2>Paragraph pages</h2>

          <ol>
            {paragraphs.map((paragraph, index) => (
              <li key={index}>
                <div className="paragraph-heading">
                  <strong>Paragraph {index + 1}</strong>
                  <span>
                    {formatPageLabel(paragraph.pages)}
                    {paragraph.pages.length > 1 && (
                      <span className="continuation"> — Continuation detected</span>
                    )}
                  </span>
                </div>
                <div className="paragraph-text">{paragraph.text}</div>
              </li>
            ))}
          </ol>
        </section>
      )}
    </main>
  );
};

Office.onReady(() => {
  const container = document.getElementById("container");

  if (!container) {
    throw new Error("Could not find the React container.");
  }

  createRoot(container).render(<App />);
});
