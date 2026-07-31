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

interface KeepTogetherTarget {
  ooxml: string;
  pageCount: number;
}

export interface KeepTogetherUpdate extends MoveParagraphsResult {
  paragraphs: string[];
}

const hasKeepLines = (ooxml: string): boolean => /<w:keepLines(?:\s[^>]*)?\/?>/i.test(ooxml);

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
