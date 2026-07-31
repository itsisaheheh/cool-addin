const CONTINUATION_SUFFIX = " (Cont'd)";
const NUMERIC_HEADING_PATTERN = /^(\d+(?:\.\d+)*)(?:\.)?(?:\s+|$)/;

export const CONTINUATION_SUFFIX_PATTERN = /\s+\(Cont['’]d\)\s*$/i;

export function parseNumericHeading(text: string): { key: string; level: number } | null {
  const trimmed = text.trim();
  if (CONTINUATION_SUFFIX_PATTERN.test(trimmed) || /^\([a-zivxlcdm]+\)/i.test(trimmed)) {
    return null;
  }

  const match = trimmed.match(NUMERIC_HEADING_PATTERN);
  if (!match) return null;

  const key = match[1];
  return { key, level: key.split(".").length };
}

export const continuationText = (headingText: string): string =>
  `${headingText.trim()}${CONTINUATION_SUFFIX}`;

export const startsWithNumericHeading = (text: string): boolean =>
  NUMERIC_HEADING_PATTERN.test(text);
