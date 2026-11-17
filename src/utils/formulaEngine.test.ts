import test from 'node:test';
import assert from 'node:assert/strict';
import {
  evaluateSheet,
  extractReferencedCells,
  gradeFormulaAnswer,
} from './formulaEngine';
import type { CellData, TestQuestion } from '../types/excel';

test('evaluates basic arithmetic, cell locks ($), SUM, AVERAGE, MIN, MAX, COUNT', () => {
  const cells: Record<string, CellData> = {
    A1: { raw: '100' },
    A2: { raw: '200' },
    A3: { raw: '300' },
    A4: { raw: 'Text' },
    B1: { raw: '0.1' },
    C1: { raw: '=SUM($A$1:$A$3)' },
    C2: { raw: '=AVERAGE(A1:A3)' },
    C3: { raw: '=MIN(A1:A3)' },
    C4: { raw: '=MAX(A1:A3)' },
    C5: { raw: '=COUNT(A1:A4)' },
    C6: { raw: '=COUNTA(A1:A4)' },
    C7: { raw: '=A2*(1-$B$1)' },
  };
  const res = evaluateSheet(cells);
  assert.equal(res.C1.value, 600);
  assert.equal(res.C2.value, 200);
  assert.equal(res.C3.value, 100);
  assert.equal(res.C4.value, 300);
  assert.equal(res.C5.value, 3);
  assert.equal(res.C6.value, 4);
  assert.equal(res.C7.value, 180);
});

test('evaluates IF, IFS, AND, OR, IFERROR, SWITCH', () => {
  const cells: Record<string, CellData> = {
    A1: { raw: '6200' },
    B1: { raw: 'Gold' },
    C1: { raw: '=IF(AND(A1>5000, B1="Gold"), A1*0.88, A1)' },
    C2: { raw: '=IFS(A1>10000, "Tier 1", A1>5000, "Tier 2", TRUE, "Tier 3")' },
    C3: { raw: '=IFERROR(10/0, "Zero Guard")' },
    C4: { raw: '=SWITCH(B1, "Silver", 0.05, "Gold", 0.12, 0)' },
    C5: { raw: '=OR(A1<100, B1="Gold")' },
  };
  const res = evaluateSheet(cells);
  assert.equal(res.C1.value, 5456);
  assert.equal(res.C2.value, 'Tier 2');
  assert.equal(res.C3.value, 'Zero Guard');
  assert.equal(res.C4.value, 0.12);
  assert.equal(res.C5.value, true);
});

test('evaluates VLOOKUP, XLOOKUP, INDEX/MATCH, SUMIFS, COUNTIFS', () => {
  const cells: Record<string, CellData> = {
    A2: { raw: 'SKU-101' }, B2: { raw: 'North' }, C2: { raw: '1200' },
    A3: { raw: 'SKU-102' }, B3: { raw: 'South' }, C3: { raw: '2400' },
    A4: { raw: 'SKU-103' }, B4: { raw: 'North' }, C4: { raw: '1800' },
    E2: { raw: 'SKU-102' },
    F2: { raw: '=VLOOKUP(E2, $A$2:$C$4, 3, FALSE)' },
    F3: { raw: '=XLOOKUP(E2, $A$2:$A$4, $C$2:$C$4, "Missing")' },
    F4: { raw: '=INDEX($C$2:$C$4, MATCH("SKU-103", $A$2:$A$4, 0))' },
    F5: { raw: '=SUMIFS($C$2:$C$4, $B$2:$B$4, "North", $C$2:$C$4, ">1500")' },
    F6: { raw: '=COUNTIFS($B$2:$B$4, "North")' },
  };
  const res = evaluateSheet(cells);
  assert.equal(res.F2.value, 2400);
  assert.equal(res.F3.value, 2400);
  assert.equal(res.F4.value, 1800);
  assert.equal(res.F5.value, 1800);
  assert.equal(res.F6.value, 2);
});

test('evaluates dynamic arrays FILTER, UNIQUE, SORT with spill', () => {
  const cells: Record<string, CellData> = {
    A2: { raw: 'East' }, B2: { raw: '400' },
    A3: { raw: 'West' }, B3: { raw: '900' },
    A4: { raw: 'East' }, B4: { raw: '700' },
    D2: { raw: '=UNIQUE(A2:A4)' },
    F2: { raw: '=FILTER(B2:B4, A2:A4="East")' },
  };
  const res = evaluateSheet(cells);
  assert.equal(res.D2.value, 'East');
  assert.equal(res.D3.value, 'West');
  assert.equal(res.D3.isSpilled, true);
  assert.equal(res.F2.value, 400);
  assert.equal(res.F3.value, 700);
});

test('evaluates text and date functions', () => {
  const cells: Record<string, CellData> = {
    A1: { raw: '  Inv-2026-09  ' },
    B1: { raw: '=TRIM(A1)' },
    B2: { raw: '=LEFT(B1, 3)' },
    B3: { raw: '=RIGHT(B1, 2)' },
    B4: { raw: '=MID(B1, 5, 4)' },
    B5: { raw: '=CONCAT(B2, "-", B4)' },
    D1: { raw: '2026-01-15' },
    D2: { raw: '=EOMONTH(D1, 1)' },
    D3: { raw: '=YEAR(D1)' },
  };
  const res = evaluateSheet(cells);
  assert.equal(res.B1.value, 'Inv-2026-09');
  assert.equal(res.B2.value, 'Inv');
  assert.equal(res.B3.value, '09');
  assert.equal(res.B4.value, '2026');
  assert.equal(res.B5.value, 'Inv-2026');
  assert.equal(res.D2.value, '2026-02-28');
  assert.equal(res.D3.value, 2026);
});

test('evaluates financial modeling functions PMT, NPV, XNPV, XIRR, LET, LAMBDA', () => {
  const cells: Record<string, CellData> = {
    A1: { raw: '-10000' }, B1: { raw: '2026-01-01' },
    A2: { raw: '4000' }, B2: { raw: '2026-07-01' },
    A3: { raw: '8000' }, B3: { raw: '2027-01-01' },
    C1: { raw: '=ROUND(PMT(0.08/12, 12, 12000), 2)' },
    C2: { raw: '=ROUND(XNPV(0.1, A1:A3, B1:B3), 2)' },
    C3: { raw: '=ROUND(XIRR(A1:A3, B1:B3), 4)' },
    C4: { raw: '=LET(rev, 5000, cost, 3200, (rev-cost)/rev)' },
    C5: { raw: '=LAMBDA(x, y, x*y+10)(5, 4)' },
  };
  const res = evaluateSheet(cells);
  assert.equal(res.C1.value, -1043.86);
  assert.ok(typeof res.C2.value === 'number' && res.C2.value > 1000);
  assert.ok(typeof res.C3.value === 'number' && res.C3.value > 0.2);
  assert.equal(res.C4.value, 0.36);
  assert.equal(res.C5.value, 30);
});

test('extractReferencedCells expands ranges and ignores string literals', () => {
  const refs = extractReferencedCells('=SUMIFS($C$2:$C$4, $B$2:$B$4, "A1:A9", $D$1)');
  assert.deepEqual(refs.sort(), ['B2', 'B3', 'B4', 'C2', 'C3', 'C4', 'D1']);
});

test('detects circular references cleanly', () => {
  const res = evaluateSheet({
    A1: { raw: '=B1+1' },
    B1: { raw: '=A1+1' },
  });
  assert.equal(res.A1.error, '#CIRC!');
});

test('gradeFormulaAnswer enforces formula prefix and avoids hardcoded constants', () => {
  const question: TestQuestion = {
    id: 'q1',
    skillLevel: 'foundation',
    title: 'Discounted Payable',
    requirement: 'Write a formula in F2 for net payable after a 12% discount if order > 5000.',
    targetCell: 'F2',
    points: 25,
    cells: {
      D2: { raw: '6000' },
      E2: { raw: '0.12' },
    },
    expectedValue: 5280,
    forbiddenHardcodedNumbers: [5280, 0.12],
    requiredPatterns: [{ regex: 'IF\\(', message: 'Use an IF() condition so orders <= 5000 do not get discounted.' }],
    optimalFormula: '=IF(D2>5000, D2*(1-$E$2), D2)',
    mentorExplanation: 'Reference E2 instead of typing 0.12.',
    hint: 'Use =IF(D2>5000, ...)',
  };

  const hardcoded = gradeFormulaAnswer('5280', question, question.cells);
  assert.equal(hardcoded.passed, false);

  const missingRef = gradeFormulaAnswer('=IF(D2>5000, D2*(1-0.12), D2)', question, question.cells);
  assert.equal(missingRef.passed, false);
  assert.ok(missingRef.bestPracticeWarnings.length > 0);

  const clean = gradeFormulaAnswer('=IF(D2>5000, D2*(1-$E$2), D2)', question, question.cells);
  assert.equal(clean.passed, true);
  assert.equal(clean.scoreAwarded, 25);
});
