export interface MoveParagraphsResult {
  paragraphsUpdated: number;
  splitParagraphsFound: number;
}

export interface RemoveMoveResult {
  paragraphsUpdated: number;
  pageBreaksRemoved: number;
}

export interface DocumentKeepLinesResult {
  paragraphsFound: number;
  paragraphsChanged: number;
  paragraphsAlreadyFormatted: number;
}

export interface DocumentKeepLinesUpdate {
  ooxml: string;
  result: DocumentKeepLinesResult;
}

export interface PaginatedParagraph {
  text: string;
  ooxml: string;
  pageCount: number;
}

export interface SplitParagraphChainUpdate {
  paragraphs: string[];
  changedIndices: number[];
  headingIndices: number[];
  bodyIndices: number[];
  result: DocumentKeepLinesResult;
}

export interface KeepPaginationSnapshot {
  paragraphs: PaginatedParagraph[];
  applyParagraph: (index: number, ooxml: string) => Promise<void>;
}

export interface KeepPaginationResult {
  paragraphsChecked: number;
  splitParagraphsFixed: number;
  paginationPasses: number;
  unfixableParagraphs: number;
}

interface KeepTogetherTarget {
  ooxml: string;
  pageCount: number;
}

export interface KeepTogetherUpdate extends MoveParagraphsResult {
  paragraphs: string[];
}

const hasKeepLines = (ooxml: string): boolean => /<w:keepLines(?:\s[^>]*)?\/?>/i.test(ooxml);
const hasKeepNext = (ooxml: string): boolean => /<w:keepNext(?:\s[^>]*)?\/?>/i.test(ooxml);

const hasPageBreakBefore = (ooxml: string): boolean =>
  /<w:pageBreakBefore(?:\s[^>]*)?\/?>/i.test(ooxml);

const addParagraphProperties = (ooxml: string, properties: string): string => {
  if (/<w:pPr(?:\s[^>]*)?>/i.test(ooxml)) {
    return ooxml.replace(/(<w:pPr(?:\s[^>]*)?>)/i, `$1${properties}`);
  }
  return ooxml.replace(/(<w:p(?:\s[^>]*)?>)/i, `$1<w:pPr>${properties}</w:pPr>`);
};

export function addMoveToNextPage(ooxml: string): string {
  const properties = `${hasPageBreakBefore(ooxml) ? "" : "<w:pageBreakBefore/>"}${
    hasKeepLines(ooxml) ? "" : "<w:keepLines/>"
  }`;

  return properties === "" ? ooxml : addParagraphProperties(ooxml, properties);
}

export function removeMoveToNextPage(ooxml: string): string {
  return ooxml.replace(
    /<w:pageBreakBefore(?:\s[^>]*)?\/>|<w:pageBreakBefore(?:\s[^>]*)?>[\s\S]*?<\/w:pageBreakBefore>/gi,
    ""
  );
}

export function enableKeepTogether(targets: KeepTogetherTarget[]): KeepTogetherUpdate {
  let splitParagraphsFound = 0;

  for (const target of targets) {
    if (target.pageCount > 1) {
      splitParagraphsFound += 1;
    }
  }

  return {
    paragraphsUpdated: targets.length,
    splitParagraphsFound,
    paragraphs: targets.map(({ ooxml }) => addMoveToNextPage(ooxml)),
  };
}

export function disableMoveToNextPage(ooxmlParagraphs: string[]): {
  result: RemoveMoveResult;
  paragraphs: string[];
} {
  const pageBreaksRemoved = ooxmlParagraphs.filter(hasPageBreakBefore).length;
  return {
    result: {
      paragraphsUpdated: ooxmlParagraphs.length,
      pageBreaksRemoved,
    },
    paragraphs: ooxmlParagraphs.map(removeMoveToNextPage),
  };
}

const PARAGRAPH_WITH_PROPERTIES_PATTERN =
  /(<w:p(?:\s[^>]*)?>)(\s*)(<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr(?:\s[^>]*)?\/>)?/gi;
const COMPLETE_PARAGRAPH_PATTERN = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gi;

const decodeXmlText = (value: string): string =>
  value
    .replace(/&#x([0-9a-f]+);/gi, (_match, code: string) =>
      String.fromCharCode(Number.parseInt(code, 16))
    )
    .replace(/&#([0-9]+);/g, (_match, code: string) =>
      String.fromCharCode(Number.parseInt(code, 10))
    )
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

const paragraphText = (ooxml: string): string => {
  const textPattern = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/gi;
  const parts: string[] = [];
  let match: RegExpExecArray | null;

  while ((match = textPattern.exec(ooxml)) !== null) {
    parts.push(decodeXmlText(match[1]));
  }

  return parts.join("").replace(/\s+/g, " ").trim();
};

const paragraphStyle = (ooxml: string): string => {
  const match = /<w:pStyle(?:\s[^>]*)?\sw:val=(?:"([^"]*)"|'([^']*)')[^>]*\/?>/i.exec(ooxml);
  return match ? (match[1] ?? match[2] ?? "") : "";
};

const hasEnabledBold = (ooxml: string): boolean => {
  const boldPattern = /<w:b(?:\s[^>]*)?\/?>/gi;
  let match: RegExpExecArray | null;

  while ((match = boldPattern.exec(ooxml)) !== null) {
    if (!/\sw:val=(?:"(?:0|false|off)"|'(?:0|false|off)')/i.test(match[0])) return true;
  }

  return false;
};

const isRecognisedHeading = (ooxml: string): boolean => {
  const text = paragraphText(ooxml);
  if (!text || text.length > 160) return false;

  const style = paragraphStyle(ooxml);
  const hasHeadingStyle = /^(?:heading|title|subtitle)/i.test(style);
  const isNumberedHeading = /^\d+(?:\.\d+)*\.?\s+\S/.test(text);
  const isLetteredHeading = /^\([a-z]\)\s+\S/i.test(text);
  const letters = text.replace(/[^A-Za-zÀ-ÖØ-öø-ÿ]/g, "");
  const isUppercaseReportHeading =
    letters.length >= 3 && letters === letters.toUpperCase() && text.length <= 120;
  const isShortBoldHeading =
    hasEnabledBold(ooxml) &&
    text.length <= 120 &&
    text.split(/\s+/).length <= 16 &&
    !/[.!?;:]$/.test(text);

  return (
    hasHeadingStyle ||
    isNumberedHeading ||
    isLetteredHeading ||
    isUppercaseReportHeading ||
    isShortBoldHeading
  );
};

export function addKeepLinesToAllParagraphs(ooxml: string): DocumentKeepLinesUpdate {
  let paragraphsFound = 0;
  let paragraphsChanged = 0;
  let paragraphsAlreadyFormatted = 0;

  const updatedOoxml = ooxml.replace(
    PARAGRAPH_WITH_PROPERTIES_PATTERN,
    (_match, paragraphStart: string, whitespace: string, paragraphProperties?: string) => {
      paragraphsFound += 1;

      if (paragraphProperties && hasKeepLines(paragraphProperties)) {
        paragraphsAlreadyFormatted += 1;
        return `${paragraphStart}${whitespace}${paragraphProperties}`;
      }

      paragraphsChanged += 1;
      if (!paragraphProperties) {
        return `${paragraphStart}${whitespace}<w:pPr><w:keepLines/></w:pPr>`;
      }
      if (/\/>$/.test(paragraphProperties)) {
        return `${paragraphStart}${whitespace}${paragraphProperties.replace(
          /\/>$/,
          "><w:keepLines/></w:pPr>"
        )}`;
      }
      return `${paragraphStart}${whitespace}${paragraphProperties.replace(
        /<\/w:pPr>$/i,
        "<w:keepLines/></w:pPr>"
      )}`;
    }
  );

  const updatedWithHeadingPagination = updatedOoxml.replace(
    COMPLETE_PARAGRAPH_PATTERN,
    (paragraph: string) =>
      isRecognisedHeading(paragraph) ? addKeepNextToParagraphOoxml(paragraph) : paragraph
  );

  return {
    ooxml: updatedWithHeadingPagination,
    result: {
      paragraphsFound,
      paragraphsChanged,
      paragraphsAlreadyFormatted,
    },
  };
}

export function addKeepLinesOnlyToAllParagraphs(ooxml: string): DocumentKeepLinesUpdate {
  let paragraphsFound = 0;
  let paragraphsChanged = 0;
  let paragraphsAlreadyFormatted = 0;

  const updatedOoxml = ooxml.replace(
    PARAGRAPH_WITH_PROPERTIES_PATTERN,
    (_match, paragraphStart: string, whitespace: string, paragraphProperties?: string) => {
      paragraphsFound += 1;

      if (paragraphProperties && hasKeepLines(paragraphProperties)) {
        paragraphsAlreadyFormatted += 1;
        return `${paragraphStart}${whitespace}${paragraphProperties}`;
      }

      paragraphsChanged += 1;
      if (!paragraphProperties) {
        return `${paragraphStart}${whitespace}<w:pPr><w:keepLines/></w:pPr>`;
      }
      if (/\/>$/.test(paragraphProperties)) {
        return `${paragraphStart}${whitespace}${paragraphProperties.replace(
          /\/>$/,
          "><w:keepLines/></w:pPr>"
        )}`;
      }
      return `${paragraphStart}${whitespace}${paragraphProperties.replace(
        /<\/w:pPr>$/i,
        "<w:keepLines/></w:pPr>"
      )}`;
    }
  );

  return {
    ooxml: updatedOoxml,
    result: {
      paragraphsFound,
      paragraphsChanged,
      paragraphsAlreadyFormatted,
    },
  };
}

export function removeKeepLinesFromAllParagraphs(ooxml: string): DocumentKeepLinesUpdate {
  let paragraphsFound = 0;
  let paragraphsChanged = 0;

  const updatedOoxml = ooxml.replace(
    PARAGRAPH_WITH_PROPERTIES_PATTERN,
    (_match, paragraphStart: string, whitespace: string, paragraphProperties?: string) => {
      paragraphsFound += 1;
      if (!paragraphProperties || !hasKeepLines(paragraphProperties)) {
        return `${paragraphStart}${whitespace}${paragraphProperties ?? ""}`;
      }

      paragraphsChanged += 1;
      return `${paragraphStart}${whitespace}${paragraphProperties.replace(
        /<w:keepLines(?:\s[^>]*)?\/>|<w:keepLines(?:\s[^>]*)?>[\s\S]*?<\/w:keepLines>/gi,
        ""
      )}`;
    }
  );

  return {
    ooxml: updatedOoxml,
    result: {
      paragraphsFound,
      paragraphsChanged,
      paragraphsAlreadyFormatted: paragraphsFound - paragraphsChanged,
    },
  };
}

export function addKeepNextToParagraphOoxml(ooxml: string): string {
  if (hasKeepNext(ooxml)) return ooxml;

  if (/<w:keepLines(?:\s[^>]*)?\/?>/i.test(ooxml)) {
    return ooxml.replace(/(<w:keepLines(?:\s[^>]*)?\/?>)/i, "<w:keepNext/>$1");
  }
  if (/<w:pPr(?:\s[^>]*)?\/>/i.test(ooxml)) {
    return ooxml.replace(
      /<w:pPr(\s[^>]*)?\/>/i,
      (_match, attributes: string | undefined) => `<w:pPr${attributes ?? ""}><w:keepNext/></w:pPr>`
    );
  }
  if (/<w:pPr(?:\s[^>]*)?>/i.test(ooxml)) {
    return ooxml.replace(/(<w:pPr(?:\s[^>]*)?>)/i, "$1<w:keepNext/>");
  }
  return ooxml.replace(/(<w:p(?:\s[^>]*)?>)/i, "$1<w:pPr><w:keepNext/></w:pPr>");
}

const addKeepLinesToParagraphOoxml = (ooxml: string): string => {
  if (hasKeepLines(ooxml)) return ooxml;
  if (/<w:pPr(?:\s[^>]*)?\/>/i.test(ooxml)) {
    return ooxml.replace(
      /<w:pPr(\s[^>]*)?\/>/i,
      (_match, attributes: string | undefined) => `<w:pPr${attributes ?? ""}><w:keepLines/></w:pPr>`
    );
  }
  return addParagraphProperties(ooxml, "<w:keepLines/>");
};

export const addKeepLinesToParagraphOnly = (ooxml: string): string =>
  addKeepLinesToParagraphOoxml(ooxml);

const NUMBERED_NOTE_HEADING_PATTERN = /^\d+(?:(?:\.\d+)+|\.)\s+\S.+$/;
const LETTERED_SUBSECTION_HEADING_PATTERN = /^\([a-z]\)\s+\S.+$/i;

const normalizedParagraphText = (text: string): string => text.replace(/\s+/g, " ").trim();

export const isNumberedNoteHeading = (text: string): boolean => {
  const normalized = normalizedParagraphText(text);
  return normalized.length <= 160 && NUMBERED_NOTE_HEADING_PATTERN.test(normalized);
};

export const isLetteredSubsectionHeading = (text: string): boolean => {
  const normalized = normalizedParagraphText(text);
  return normalized.length <= 160 && LETTERED_SUBSECTION_HEADING_PATTERN.test(normalized);
};

export const isNonNumberedReportHeading = (paragraph: PaginatedParagraph): boolean => {
  const text = normalizedParagraphText(paragraph.text);
  if (!text || text.length > 160) return false;

  const letters = text.replace(/[^A-Za-zÀ-ÖØ-öø-ÿ]/g, "");
  const isUppercase = letters.length >= 3 && letters === letters.toUpperCase();
  const style = paragraphStyle(paragraph.ooxml);
  return isUppercase || /^(?:heading|title|subtitle)/i.test(style);
};

const isAnyHeading = (paragraph: PaginatedParagraph): boolean =>
  isNumberedNoteHeading(paragraph.text) ||
  isLetteredSubsectionHeading(paragraph.text) ||
  isNonNumberedReportHeading(paragraph);

const previousNonEmptyIndex = (
  paragraphs: PaginatedParagraph[],
  startIndex: number
): number | null => {
  for (let index = startIndex - 1; index >= 0; index -= 1) {
    if (normalizedParagraphText(paragraphs[index].text) !== "") return index;
  }
  return null;
};

export function formatSplitParagraphChains(
  paragraphs: PaginatedParagraph[]
): SplitParagraphChainUpdate {
  const updatedParagraphs = paragraphs.map(({ ooxml }) => ooxml);
  const applicableIndices = new Set<number>();
  const headingIndices = new Set<number>();
  const bodyIndices = new Set<number>();

  paragraphs.forEach((paragraph, bodyIndex) => {
    if (paragraph.pageCount <= 1 || isAnyHeading(paragraph)) {
      return;
    }

    applicableIndices.add(bodyIndex);
    bodyIndices.add(bodyIndex);
    updatedParagraphs[bodyIndex] = addKeepLinesToParagraphOoxml(updatedParagraphs[bodyIndex]);

    const nearestIndex = previousNonEmptyIndex(paragraphs, bodyIndex);
    if (nearestIndex === null) return;

    if (isLetteredSubsectionHeading(paragraphs[nearestIndex].text)) {
      applicableIndices.add(nearestIndex);
      headingIndices.add(nearestIndex);
      updatedParagraphs[nearestIndex] = addKeepNextToParagraphOoxml(
        updatedParagraphs[nearestIndex]
      );
    }
  });

  const changedIndices = Array.from(applicableIndices).filter(
    (index) => updatedParagraphs[index] !== paragraphs[index].ooxml
  );
  changedIndices.sort((left, right) => left - right);

  return {
    paragraphs: updatedParagraphs,
    changedIndices,
    headingIndices: Array.from(headingIndices).sort((left, right) => left - right),
    bodyIndices: Array.from(bodyIndices).sort((left, right) => left - right),
    result: {
      paragraphsFound: paragraphs.length,
      paragraphsChanged: changedIndices.length,
      paragraphsAlreadyFormatted: applicableIndices.size - changedIndices.length,
    },
  };
}

export async function validateKeepLinesPagination(
  totalParagraphCount: number,
  scan: () => Promise<KeepPaginationSnapshot>
): Promise<KeepPaginationResult> {
  const maximumPasses = Math.max(1, totalParagraphCount * 2 + 1);
  const unfixableIndices = new Set<number>();
  const formattedIndices = new Set<number>();
  let paragraphsChecked = totalParagraphCount;
  let paginationPasses = 0;

  while (paginationPasses < maximumPasses) {
    paginationPasses += 1;
    const snapshot = await scan();
    paragraphsChecked = snapshot.paragraphs.length;

    const splitIndex = snapshot.paragraphs.findIndex(
      (paragraph, index) => paragraph.pageCount > 1 && !unfixableIndices.has(index)
    );
    if (splitIndex < 0) break;

    const splitParagraph = snapshot.paragraphs[splitIndex];
    if (hasKeepLines(splitParagraph.ooxml)) {
      unfixableIndices.add(splitIndex);
      continue;
    }

    await snapshot.applyParagraph(splitIndex, addKeepLinesToParagraphOoxml(splitParagraph.ooxml));
    formattedIndices.add(splitIndex);
  }

  return {
    paragraphsChecked,
    splitParagraphsFixed: Array.from(formattedIndices).filter(
      (index) => !unfixableIndices.has(index)
    ).length,
    paginationPasses,
    unfixableParagraphs: unfixableIndices.size,
  };
}
