import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";
import {
  ParagraphPageResult,
  removeContinuationMarkers,
  removeKeepAllParagraphsTogether,
} from "./word";
import {
  runAddContdHeadingsOnly,
  runCheckDocumentOnly,
  runKeepParagraphsIntactOnly,
} from "./feature-actions";

const formatPageLabel = (pages: number[]): string => {
  if (pages.length === 0) return "Page unavailable";
  if (pages.length === 1) return `Page ${pages[0]}`;
  return `Pages ${pages.join("–")}`;
};

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : JSON.stringify(error);

const App = (): React.ReactElement => {
  const [checkStatus, setCheckStatus] = React.useState("Ready to scan.");
  const [keepStatus, setKeepStatus] = React.useState("Ready.");
  const [contdStatus, setContdStatus] = React.useState("Ready.");
  const [pageCount, setPageCount] = React.useState<number | null>(null);
  const [paragraphs, setParagraphs] = React.useState<ParagraphPageResult[]>([]);
  const [isChecking, setIsChecking] = React.useState(false);
  const [isKeepingParagraphs, setIsKeepingParagraphs] = React.useState(false);
  const [isAddingContd, setIsAddingContd] = React.useState(false);
  const isBusy = isChecking || isKeepingParagraphs || isAddingContd;

  const handleCheckDocument = async (): Promise<void> => {
    setCheckStatus("Scanning and reporting document issues...");
    setIsChecking(true);
    setPageCount(null);
    setParagraphs([]);

    try {
      const result = await runCheckDocumentOnly();
      setPageCount(result.pagination.pageCount);
      setParagraphs(result.pagination.paragraphs);
      setCheckStatus(
        `Check completed. Found ${result.continuation.continuingSectionsFound} continuing sections across ${result.continuation.continuationPagesFound} continuation pages and ${result.continuation.duplicatesSkipped} existing duplicate continuation headings. No document changes were made.`
      );
    } catch (error) {
      console.error("Check Document error:", error);
      setCheckStatus(`Check failed: ${errorMessage(error)}`);
    } finally {
      setIsChecking(false);
    }
  };

  const handleKeepParagraphsIntact = async (): Promise<void> => {
    setKeepStatus("Applying paragraph pagination formatting...");
    setIsKeepingParagraphs(true);

    try {
      const result = await runKeepParagraphsIntactOnly();
      setKeepStatus(
        `Completed. ${result.paragraphsFound.toLocaleString()} paragraphs checked. Split paragraphs fixed: ${result.splitParagraphsFixed}. Orphan headings fixed: ${result.orphanHeadingsFixed}. ${result.paginationPasses} pagination passes completed. ${result.unfixableParagraphs} oversized paragraphs remain split.`
      );
    } catch (error) {
      console.error("Keep Paragraphs Intact error:", error);
      setKeepStatus(`Keep Paragraphs Intact failed: ${errorMessage(error)}`);
    } finally {
      setIsKeepingParagraphs(false);
    }
  };

  const handleAddContdHeadings = async (): Promise<void> => {
    setContdStatus("Detecting continuation pages and adding CONT’D headings...");
    setIsAddingContd(true);

    try {
      const result = await runAddContdHeadingsOnly();
      setContdStatus(
        `Completed. Inserted ${result.headingsInserted} CONT’D headings across ${result.continuationPagesFound} continuation pages; skipped ${result.duplicatesSkipped} duplicates. ${result.limitationMessage}`
      );
    } catch (error) {
      console.error("Add CONT’D Headings error:", error);
      setContdStatus(`Add CONT’D Headings failed: ${errorMessage(error)}`);
    } finally {
      setIsAddingContd(false);
    }
  };

  const handleUndoContdHeadings = async (): Promise<void> => {
    setContdStatus("Removing continuation headings...");
    setIsAddingContd(true);

    try {
      const removedCount = await removeContinuationMarkers();
      setContdStatus(`Removed ${removedCount} continuation headings or legacy markers.`);
    } catch (error) {
      console.error("Remove CONT’D Headings error:", error);
      setContdStatus(`Remove CONT’D Headings failed: ${errorMessage(error)}`);
    } finally {
      setIsAddingContd(false);
    }
  };

  const handleUndoKeepParagraphs = async (): Promise<void> => {
    setKeepStatus("Removing paragraph pagination formatting...");
    setIsKeepingParagraphs(true);

    try {
      const result = await removeKeepAllParagraphsTogether();
      setKeepStatus(
        `Removed Keep lines together from ${result.paragraphsChanged} of ${result.paragraphsFound} paragraphs. Other formatting was preserved.`
      );
    } catch (error) {
      console.error("Undo Keep Paragraphs Intact error:", error);
      setKeepStatus(`Undo Keep Paragraphs Intact failed: ${errorMessage(error)}`);
    } finally {
      setIsKeepingParagraphs(false);
    }
  };

  return (
    <main className="app">
      <h1>Document Checker</h1>

      <section className="feature-card">
        <h2>Check Document</h2>
        <p>Scans the report and reports pagination and continuation issues without changing it.</p>
        <button type="button" onClick={handleCheckDocument} disabled={isBusy}>
          {isChecking ? "Checking..." : "Check Document"}
        </button>
        <p className="feature-status">
          <strong>Status:</strong> {checkStatus}
        </p>
      </section>

      <section className="feature-card">
        <h2>Keep Paragraphs Intact</h2>
        <p>
          Prevents paragraphs and their immediate headings from splitting awkwardly across pages.
        </p>
        <div className="button-row">
          <button type="button" onClick={handleKeepParagraphsIntact} disabled={isBusy}>
            {isKeepingParagraphs ? "Formatting..." : "Keep Paragraphs Intact"}
          </button>
          <button
            className="undo-button"
            type="button"
            onClick={handleUndoKeepParagraphs}
            disabled={isBusy}
            aria-label="Undo Keep Paragraphs Intact"
            title="Undo Keep Paragraphs Intact"
          >
            {"\u21B6"}
          </button>
        </div>
        <p className="feature-status">
          <strong>Status:</strong> {keepStatus}
        </p>
      </section>

      <section className="feature-card">
        <h2>Add CONT’D Headings</h2>
        <p>Adds continuation headings where the existing CONT’D rules determine they are needed.</p>
        <div className="button-row">
          <button type="button" onClick={handleAddContdHeadings} disabled={isBusy}>
            {isAddingContd ? "Working..." : "Add CONT’D Headings"}
          </button>
          <button
            className="undo-button"
            type="button"
            onClick={handleUndoContdHeadings}
            disabled={isBusy}
            aria-label="Undo CONT’D headings"
            title="Undo CONT’D headings"
          >
            {"\u21B6"}
          </button>
        </div>
        <p className="feature-status">
          <strong>Status:</strong> {contdStatus}
        </p>
      </section>

      {pageCount !== null && (
        <p className="page-summary">
          <strong>Total pages:</strong> {pageCount}
        </p>
      )}

      {paragraphs.length > 0 && (
        <section className="results">
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
  if (!container) throw new Error("Could not find the React container.");
  createRoot(container).render(<App />);
});
