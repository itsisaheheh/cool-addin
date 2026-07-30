import * as React from "react";
import { createRoot } from "react-dom/client";
import "./taskpane.css";

const App = (): React.ReactElement => {
  const [status, setStatus] = React.useState("Ready");
  const [paragraphs, setParagraphs] = React.useState<string[]>([]);

  const checkDocument = async (): Promise<void> => {
    setStatus("Checking document...");

    try {
      await Word.run(async (context) => {
        const documentParagraphs = context.document.body.paragraphs;

        documentParagraphs.load("items/text");
        await context.sync();

        const paragraphTexts = documentParagraphs.items
          .map((paragraph) => paragraph.text)
          .filter((text) => text.trim() !== "");

        setParagraphs(paragraphTexts);
        setStatus(`Success! Found ${paragraphTexts.length} paragraphs.`);

        console.log("Paragraphs:", paragraphTexts);
      });
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

      {paragraphs.length > 0 && (
        <section>
          <h2>Document paragraphs</h2>

          <ol>
            {paragraphs.map((paragraph, index) => (
              <li key={index}>
  <strong>Paragraph {index + 1}</strong>
  <br />
  {paragraph}
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