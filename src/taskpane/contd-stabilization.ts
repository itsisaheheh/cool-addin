export interface ContdPassResult {
  continuingSectionsFound: number;
  continuationPagesFound: number;
  headingsInserted: number;
  duplicatesSkipped: number;
  limitationMessage: string;
  paginationChanged?: boolean;
}

export interface ContdStabilizationResult extends ContdPassResult {
  passesCompleted: number;
  reachedSafetyLimit: boolean;
}

export async function runContdInsertionUntilStable(
  maximumPasses: number,
  runPass: (passNumber: number) => Promise<ContdPassResult>
): Promise<ContdStabilizationResult> {
  const safeMaximum = Math.max(1, maximumPasses);
  let passesCompleted = 0;
  let continuingSectionsFound = 0;
  let continuationPagesFound = 0;
  let headingsInserted = 0;
  let duplicatesSkipped = 0;
  let limitationMessage = "";
  let reachedSafetyLimit = true;

  while (passesCompleted < safeMaximum) {
    const result = await runPass(passesCompleted + 1);
    passesCompleted += 1;
    continuingSectionsFound = Math.max(continuingSectionsFound, result.continuingSectionsFound);
    continuationPagesFound = Math.max(continuationPagesFound, result.continuationPagesFound);
    headingsInserted += result.headingsInserted;
    duplicatesSkipped = Math.max(duplicatesSkipped, result.duplicatesSkipped);
    limitationMessage = result.limitationMessage;

    if (result.headingsInserted === 0 && !result.paginationChanged) {
      reachedSafetyLimit = false;
      break;
    }
  }

  return {
    continuingSectionsFound,
    continuationPagesFound,
    headingsInserted,
    duplicatesSkipped,
    limitationMessage: `${limitationMessage} Full-document CONT’D scanning ${
      reachedSafetyLimit
        ? `stopped at the ${safeMaximum}-pass safety limit`
        : `stabilized after ${passesCompleted} pass${passesCompleted === 1 ? "" : "es"}`
    }.`,
    passesCompleted,
    reachedSafetyLimit,
  };
}
