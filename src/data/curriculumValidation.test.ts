import test from 'node:test';
import assert from 'node:assert/strict';
import { evaluateSheet, gradeFormulaAnswer } from '../utils/formulaEngine';
import { FORMULA_VAULT } from './formulaVaultData';
import { LIVE_EXAMPLES } from './liveExamplesData';
import { PRACTICAL_TESTS } from './practicalTestsData';
import { DEFAULT_PRACTICE_PRESETS } from './practiceScenarios';

test('all 70 Formula Vault presets evaluate without errors', () => {
  assert.equal(FORMULA_VAULT.length, 70);
  for (const formula of FORMULA_VAULT) {
    const res = evaluateSheet(formula.preset.cells);
    const target = res[formula.preset.targetCell];
    assert.ok(target, `Missing target cell for ${formula.name}`);
    assert.equal(target.error, undefined, `${formula.name} errored with ${target.error}`);
  }
});

test('all 4 default Live Practice presets evaluate cleanly', () => {
  for (const [level, preset] of Object.entries(DEFAULT_PRACTICE_PRESETS)) {
    const res = evaluateSheet(preset.cells);
    for (const [coord, cell] of Object.entries(res)) {
      assert.equal(cell.error, undefined, `Default preset ${level} cell ${coord} errored: ${cell.error}`);
    }
  }
});

test('all 12 Live Examples evaluate cleanly without errors', () => {
  assert.equal(LIVE_EXAMPLES.length, 12);
  for (const ex of LIVE_EXAMPLES) {
    const res = evaluateSheet(ex.preset.cells);
    for (const [coord, cell] of Object.entries(res)) {
      assert.equal(cell.error, undefined, `Example ${ex.id} cell ${coord} errored: ${cell.error}`);
    }
  }
});

test('all 12 Practical Test optimal formulas pass grading', () => {
  assert.equal(PRACTICAL_TESTS.length, 12);
  for (const q of PRACTICAL_TESTS) {
    const grade = gradeFormulaAnswer(q.optimalFormula, q, q.cells);
    assert.equal(
      grade.passed,
      true,
      `Question ${q.id} failed optimal formula (${q.optimalFormula}): ${grade.feedbackTitle} | ${grade.feedbackDetails.join('; ')} | ${grade.bestPracticeWarnings.join('; ')}`
    );
  }
});
