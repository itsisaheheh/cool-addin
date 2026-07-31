import { parseNumericHeading, startsWithNumericHeading } from "./continuation-format";

export interface CheckParagraphInput {
  text: string;
  listPrefix: string;
}

export interface NotesNumberedHeading {
  key: string;
  title: string;
}

export interface NotesParagraphDiagnostic {
  paragraphIndex: number;
  text: string;
  listPrefix: string;
  result: string;
}

export interface NotesScanDiagnostics {
  notesHeadingFound: boolean;
  notesHeadingIndex: number | null;
  paragraphsScannedAfterNotes: number;
  firstNonEmptyParagraphs: NotesParagraphDiagnostic[];
}

export interface NotesHeadingScanResult {
  headings: NotesNumberedHeading[];
  diagnostics: NotesScanDiagnostics;
}

const NOTES_HEADING_PATTERN = /^NOTES TO THE FINANCIAL STATEMENTS$/i;

const normalizeWordText = (text: string): string =>
  text
    .normalize("NFKC")
    // Word paragraph text can contain cell markers and other non-printing controls.
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u001f\u007f-\u009f\u200b-\u200f\u202a-\u202e\u2060\ufeff]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export function inspectNumberedHeadingsInNotes(
  paragraphs: CheckParagraphInput[]
): NotesHeadingScanResult {
  let insideNotes = false;
  let notesHeadingIndex: number | null = null;
  const headings: NotesNumberedHeading[] = [];
  const firstNonEmptyParagraphs: NotesParagraphDiagnostic[] = [];

  paragraphs.forEach((paragraph, paragraphIndex) => {
    const paragraphText = normalizeWordText(paragraph.text);
    if (NOTES_HEADING_PATTERN.test(paragraphText)) {
      insideNotes = true;
      notesHeadingIndex = paragraphIndex;
      return;
    }
    if (!insideNotes) return;

    const listPrefix = normalizeWordText(paragraph.listPrefix ?? "");
    const numeric = parseNumericHeading(paragraphText) ?? parseNumericHeading(`${listPrefix} `);
    let result = "Rejected: no numeric prefix in paragraph.text or Word listString.";
    if (numeric) {
      const completeText =
        listPrefix && !startsWithNumericHeading(paragraphText)
          ? `${listPrefix} ${paragraphText}`
          : paragraphText;
      headings.push({
        key: numeric.key,
        title: completeText.replace(/^\d+(?:\.\d+)*\.?\s*/, "").trim() || "Untitled section",
      });
      result = `Accepted as section ${numeric.key}.`;
    }

    if ((paragraphText || listPrefix) && firstNonEmptyParagraphs.length < 20) {
      firstNonEmptyParagraphs.push({
        paragraphIndex,
        text: paragraphText,
        listPrefix,
        result,
      });
    }
  });

  return {
    headings,
    diagnostics: {
      notesHeadingFound: notesHeadingIndex !== null,
      notesHeadingIndex,
      paragraphsScannedAfterNotes:
        notesHeadingIndex === null ? 0 : paragraphs.length - notesHeadingIndex - 1,
      firstNonEmptyParagraphs,
    },
  };
}

export function findNumberedHeadingsInNotes(
  paragraphs: CheckParagraphInput[]
): NotesNumberedHeading[] {
  return inspectNumberedHeadingsInNotes(paragraphs).headings;
}
