import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";
import {
  analyzeDocumentPagination,
  ParagraphPageResult,
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

  const checkDocument = async (): Promise<void> => {
    setStatus("Checking document...");
    setPageCount(null);
    setParagraphs([]);

    try {
      const result = await analyzeDocumentPagination();

      setPageCount(result.pageCount);
      setParagraphs(result.paragraphs);
      setStatus(
        `Success! Found ${result.paragraphs.length} paragraphs across ${result.pageCount} pages.`
      );

      console.log("Document pagination:", result);
    } catch (error) {
      console.error("Word error:", error);

      const message =
        error instanceof Error ? error.message : JSON.stringify(error);

      setStatus(`Error: ${message}`);
    }
  };

  return (
    <main className="app">
      <h1>Document Continuation Checker</h1>

      <p>Check whether document content continues onto another page.</p>

      <button type="button" onClick={checkDocument}>
        Check Document
      </button>

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
