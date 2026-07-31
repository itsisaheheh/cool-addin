import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";
import {
  NumberedHeadingCheckResult,
  removeContinuationMarkers,
  removeKeepAllParagraphsTogether,
} from "./word";
import {
  runAddContdHeadingsOnly,
  runCheckDocumentOnly,
  runKeepParagraphsIntactOnly,
} from "./feature-actions";

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : JSON.stringify(error);

const App = (): React.ReactElement => {
  const [openDescriptions, setOpenDescriptions] = React.useState({
    checkDocument: false,
    keepParagraphs: false,
    contdHeadings: false,
  });
  const [checkStatus, setCheckStatus] = React.useState("Ready to scan.");
  const [checkResults, setCheckResults] = React.useState<NumberedHeadingCheckResult[]>([]);
  const [keepStatus, setKeepStatus] = React.useState("Ready.");
  const [contdStatus, setContdStatus] = React.useState("Ready.");
  const [isChecking, setIsChecking] = React.useState(false);
  const [isKeepingParagraphs, setIsKeepingParagraphs] = React.useState(false);
  const [isAddingContd, setIsAddingContd] = React.useState(false);
  const isBusy = isChecking || isKeepingParagraphs || isAddingContd;

  const toggleDescription = (key: keyof typeof openDescriptions): void => {
    setOpenDescriptions((previous) => ({
      ...previous,
      [key]: !previous[key],
    }));
  };

  const handleCheckDocument = async (): Promise<void> => {
    setCheckStatus("Scanning for numbered headings...");
    setIsChecking(true);
    setCheckResults([]);

    try {
      const result = await runCheckDocumentOnly();
      setCheckResults(result.numberedHeadings);
      setCheckStatus(
        `Detected numbered headings: ${result.numberedHeadings.length}. No document changes were made.`
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
        <button
          className="feature-toggle"
          type="button"
          onClick={() => toggleDescription("checkDocument")}
          aria-expanded={openDescriptions.checkDocument}
          aria-controls="check-document-description"
        >
          <span>Check Document</span>
          <span aria-hidden="true">{openDescriptions.checkDocument ? "▲" : "▼"}</span>
        </button>
        {openDescriptions.checkDocument && (
          <p id="check-document-description">
            Scans the report and reports pagination and continuation issues without changing it.
          </p>
        )}
        <button type="button" onClick={handleCheckDocument} disabled={isBusy}>
          {isChecking ? "Checking..." : "Check Document"}
        </button>
        <p className="feature-status">
          <strong>Status:</strong> {checkStatus}
        </p>
        {checkResults.length > 0 && (
          <div className="check-results">
            <p>
              <strong>Detected numbered headings: {checkResults.length}</strong>
            </p>
            <ol className="check-heading-list">
              {checkResults.map((heading, index) => (
                <li key={`${heading.key}-${index}`}>
                  <strong>{heading.key}</strong> — {heading.title}
                </li>
              ))}
            </ol>
          </div>
        )}
      </section>

            <section className="feature-card">
        <button
          className="feature-toggle"
          type="button"
          onClick={() => toggleDescription("contdHeadings")}
          aria-expanded={openDescriptions.contdHeadings}
          aria-controls="contd-headings-description"
        >
          <span>Add CONT’D Headings</span>
          <span aria-hidden="true">{openDescriptions.contdHeadings ? "▲" : "▼"}</span>
        </button>
        {openDescriptions.contdHeadings && (
          <p id="contd-headings-description">
            Adds continuation headings where the existing CONT’D rules determine they are needed.
          </p>
        )}
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

      <section className="feature-card">
        <button
          className="feature-toggle"
          type="button"
          onClick={() => toggleDescription("keepParagraphs")}
          aria-expanded={openDescriptions.keepParagraphs}
          aria-controls="keep-paragraphs-description"
        >
          <span>Keep Paragraphs Intact</span>
          <span aria-hidden="true">{openDescriptions.keepParagraphs ? "▲" : "▼"}</span>
        </button>
        {openDescriptions.keepParagraphs && (
          <p id="keep-paragraphs-description">
            Prevents paragraphs and their immediate headings from splitting awkwardly across pages.
          </p>
        )}
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
    </main>
  );
};

Office.onReady(() => {
  const container = document.getElementById("container");
  if (!container) throw new Error("Could not find the React container.");
  createRoot(container).render(<App />);
});
