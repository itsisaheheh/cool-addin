import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";
import {
  analyzeDocumentPagination,
  assessContinuationMarkers,
  ParagraphPageResult,
  removeContinuationMarkers,
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
  const [includeContinuedOnMarker, setIncludeContinuedOnMarker] = React.useState(false);
  const [isChecking, setIsChecking] = React.useState(false);

  const checkDocument = async (): Promise<void> => {
    setStatus("Checking document...");
    setIsChecking(true);
    setPageCount(null);
    setParagraphs([]);

    try {
      const insertionResult = await assessContinuationMarkers(includeContinuedOnMarker);
      const paginationResult = await analyzeDocumentPagination();

      setPageCount(paginationResult.pageCount);
      setParagraphs(paginationResult.paragraphs);
      setStatus(
        `Found ${insertionResult.boundariesFound} continuation boundaries and inserted ${insertionResult.markersInserted} markers. ${insertionResult.limitationMessage}`
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
      setStatus(`Removed ${removedCount} continuation markers.`);
    } catch (error) {
      console.error("Word error:", error);

      const message = error instanceof Error ? error.message : JSON.stringify(error);
      setStatus(`Error: ${message}`);
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
          checked={includeContinuedOnMarker}
          onChange={(event) => setIncludeContinuedOnMarker(event.target.checked)}
          disabled={isChecking}
        />
        Insert (CONT&apos;D) at the beginning of the next page
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
