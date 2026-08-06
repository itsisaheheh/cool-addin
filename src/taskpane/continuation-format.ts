const CONTINUATION_SUFFIX = " (Cont'd)";
const MAIN_CONTINUATION_SUFFIX = " (CONT'D)";
const NUMERIC_HEADING_PATTERN = /^(?:(\d+)\.\s+\S.*|(\d+\.\d+)\s+\S.*)$/;

const normalizeHeadingWhitespace = (text: string): string =>
  text
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const CONTINUATION_SUFFIX_PATTERN = /\s+\(Cont['’]d\)\s*$/i;

export const normalizeContinuationHeadingText = (text: string): string =>
  text.replace(/[’']/g, "'").replace(/\s+/g, " ").trim().toLocaleLowerCase();

export function parseNumericHeading(text: string): { key: string; level: number } | null {
  const trimmed = normalizeHeadingWhitespace(text);
  if (CONTINUATION_SUFFIX_PATTERN.test(trimmed) || /^\([a-zivxlcdm]+\)/i.test(trimmed)) {
    return null;
  }

  const match = trimmed.match(NUMERIC_HEADING_PATTERN);
  if (!match) return null;

  const key = match[1] ?? match[2];
  return { key, level: key.split(".").length };
}

export const continuationText = (headingText: string, isMainHeading = false): string =>
  `${headingText.trim()}${isMainHeading ? MAIN_CONTINUATION_SUFFIX : CONTINUATION_SUFFIX}`;

export const startsWithNumericHeading = (text: string): boolean =>
  NUMERIC_HEADING_PATTERN.test(normalizeHeadingWhitespace(text));

export const MAX_CONTINUATION_PAGINATION_PASSES = 2;

export type ContinuationPlacement =
  "prepare-and-repaginate" | "before-complete-paragraph" | "at-rendered-page-start";

export function continuationPlacement(
  startsInsideParagraph: boolean,
  preparationPass: boolean
): ContinuationPlacement {
  if (preparationPass) return "prepare-and-repaginate";
  return startsInsideParagraph ? "at-rendered-page-start" : "before-complete-paragraph";
}

export type ContinuationPageEligibility = "prepare" | "insert" | "skip";

export function pageRequiresContinuation(options: {
  activeNoteStartedEarlier: boolean;
  pageBeginsWithOriginalHeading: boolean;
}): boolean {
  return options.activeNoteStartedEarlier && !options.pageBeginsWithOriginalHeading;
}

export function validateContinuationPageTop(options: {
  anchorParagraphIndex: number;
  requiredHeadingTexts: string[];
  existingHeadings: Array<{ documentIndex: number; text: string }>;
}): { validPrefixTexts: string[]; misplacedParagraphIndex: number | null } {
  const required = options.requiredHeadingTexts.map(normalizeContinuationHeadingText);
  const existing = options.existingHeadings
    .map((heading) => ({
      ...heading,
      normalizedText: normalizeContinuationHeadingText(heading.text),
    }))
    .filter((heading) => required.includes(heading.normalizedText))
    .sort((left, right) => left.documentIndex - right.documentIndex);
  const validPrefixTexts: string[] = [];

  for (const heading of existing) {
    if (
      heading.documentIndex < options.anchorParagraphIndex &&
      heading.normalizedText === required[validPrefixTexts.length]
    ) {
      validPrefixTexts.push(heading.normalizedText);
    } else {
      return { validPrefixTexts, misplacedParagraphIndex: heading.documentIndex };
    }
  }
  return { validPrefixTexts, misplacedParagraphIndex: null };
}

export function continuationPageEligibility(options: {
  sectionStartPage: number;
  currentPage: number;
  anchorStartPage: number | null;
  anchorIsOriginalHeading: boolean;
  anchorSpansFromEarlierPage: boolean;
}): ContinuationPageEligibility {
  if (options.sectionStartPage >= options.currentPage || options.anchorIsOriginalHeading) {
    return "skip";
  }
  if (options.anchorSpansFromEarlierPage) return "prepare";
  return options.anchorStartPage === options.currentPage ? "insert" : "skip";
}

export function isOrphanOriginalHeading(options: {
  headingStartPage: number | null;
  nextContentStartPage: number | null;
  nextParagraphIsNumberedHeading: boolean;
}): boolean {
  return (
    options.headingStartPage !== null &&
    options.nextContentStartPage !== null &&
    !options.nextParagraphIsNumberedHeading &&
    options.nextContentStartPage > options.headingStartPage
  );
}
