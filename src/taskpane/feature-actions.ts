import { addContdHeadings, checkDocumentIssues, keepAllParagraphsOnOnePage } from "./word";

export const runCheckDocumentOnly = () => checkDocumentIssues();

export const runKeepParagraphsIntactOnly = () => keepAllParagraphsOnOnePage();

export const runAddContdHeadingsOnly = () => addContdHeadings();
