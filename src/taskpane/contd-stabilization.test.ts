import { runContdInsertionUntilStable } from "./contd-stabilization";

/* global describe, test, expect, jest */

const passResult = (headingsInserted: number, duplicatesSkipped = 0) => ({
  continuingSectionsFound: headingsInserted > 0 ? 1 : 0,
  continuationPagesFound: headingsInserted > 0 ? 1 : 0,
  headingsInserted,
  duplicatesSkipped,
  limitationMessage: "Existing CONT’D rules applied.",
});

describe("Add CONT'D Headings full-document stabilization", () => {
  test("continues scanning when a later heading becomes eligible after an insertion", async () => {
    const runPass = jest
      .fn()
      .mockResolvedValueOnce(passResult(1))
      .mockResolvedValueOnce(passResult(1))
      .mockResolvedValueOnce(passResult(0));

    const result = await runContdInsertionUntilStable(10, runPass);

    expect(runPass).toHaveBeenCalledTimes(3);
    expect(result.headingsInserted).toBe(2);
    expect(result.reachedSafetyLimit).toBe(false);
  });

  test("starts a fresh numbered pass after each single insertion", async () => {
    const runPass = jest
      .fn()
      .mockResolvedValueOnce(passResult(1))
      .mockResolvedValueOnce(passResult(1))
      .mockResolvedValueOnce(passResult(0));

    const result = await runContdInsertionUntilStable(10, runPass);

    expect(result.passesCompleted).toBe(3);
    expect(result.headingsInserted).toBe(2);
    expect(runPass.mock.calls).toEqual([[1], [2], [3]]);
  });

  test("stops after a complete pass inserts no missing headings", async () => {
    const runPass = jest.fn().mockResolvedValue(passResult(0));

    const result = await runContdInsertionUntilStable(10, runPass);

    expect(runPass).toHaveBeenCalledTimes(1);
    expect(result.passesCompleted).toBe(1);
  });

  test("rescans after a continuation pass changes pagination without inserting a heading", async () => {
    const runPass = jest
      .fn<Promise<ReturnType<typeof passResult> & { paginationChanged?: boolean }>, [number]>()
      .mockResolvedValueOnce({ ...passResult(0), paginationChanged: true })
      .mockResolvedValueOnce(passResult(0));

    const result = await runContdInsertionUntilStable(4, runPass);

    expect(runPass).toHaveBeenCalledTimes(2);
    expect(result.passesCompleted).toBe(2);
    expect(result.reachedSafetyLimit).toBe(false);
  });

  test("preserves duplicate-prevention results without retrying duplicates", async () => {
    const runPass = jest.fn().mockResolvedValue(passResult(0, 2));

    const result = await runContdInsertionUntilStable(10, runPass);

    expect(runPass).toHaveBeenCalledTimes(1);
    expect(result.duplicatesSkipped).toBe(2);
    expect(result.headingsInserted).toBe(0);
  });

  test("uses the maximum-pass safety limit to prevent an infinite loop", async () => {
    const runPass = jest.fn().mockResolvedValue(passResult(1));

    const result = await runContdInsertionUntilStable(3, runPass);

    expect(runPass).toHaveBeenCalledTimes(3);
    expect(result.reachedSafetyLimit).toBe(true);
  });
});
