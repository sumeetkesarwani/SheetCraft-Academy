import React, { useEffect, useMemo, useState } from 'react';
import {
  Award,
  CheckCircle2,
  Clock,
  Eye,
  HelpCircle,
  RotateCcw,
  XCircle,
} from 'lucide-react';
import type { CellData, GradeResult, SkillLevel, TestQuestion } from '../types/excel';
import { PRACTICAL_TESTS } from '../data/practicalTestsData';
import { SKILL_ORDER, SKILL_THEMES } from '../data/skillThemes';
import { gradeFormulaAnswer } from '../utils/formulaEngine';
import { SpreadsheetGrid } from './SpreadsheetGrid';

interface PracticalTestViewProps {
  activeSkillLevel: SkillLevel;
  onSkillLevelChange: (level: SkillLevel) => void;
}

export const PracticalTestView: React.FC<PracticalTestViewProps> = ({
  activeSkillLevel,
  onSkillLevelChange,
}) => {
  const theme = SKILL_THEMES[activeSkillLevel];

  const sectionQuestions = useMemo(
    () => PRACTICAL_TESTS.filter((q) => q.skillLevel === activeSkillLevel),
    [activeSkillLevel]
  );

  const [questionIndex, setQuestionIndex] = useState(0);
  const [gradesById, setGradesById] = useState<Record<string, GradeResult>>({});
  const [cellsByQuestionId, setCellsByQuestionId] = useState<Record<string, Record<string, CellData>>>({});
  const [selectedCell, setSelectedCell] = useState('F2');
  const [showHint, setShowHint] = useState(false);
  const [showSolution, setShowSolution] = useState(false);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);

  // Reset question index & select target cell when switching section exam
  useEffect(() => {
    setQuestionIndex(0);
    setShowHint(false);
    setShowSolution(false);
  }, [activeSkillLevel]);

  const currentQuestion: TestQuestion | undefined = sectionQuestions[questionIndex];

  useEffect(() => {
    if (!currentQuestion) return;
    setSelectedCell(currentQuestion.targetCell);
    setShowHint(false);
    setShowSolution(false);
  }, [currentQuestion]);

  // Live exam clock
  useEffect(() => {
    const id = window.setInterval(() => {
      setElapsedSeconds((s) => s + 1);
    }, 1000);
    return () => window.clearInterval(id);
  }, []);

  const currentCells = useMemo(() => {
    if (!currentQuestion) return {};
    return cellsByQuestionId[currentQuestion.id] || currentQuestion.cells;
  }, [cellsByQuestionId, currentQuestion]);

  const handleCellChange = (coord: string, newRaw: string) => {
    if (!currentQuestion) return;
    setCellsByQuestionId((prev) => {
      const base = prev[currentQuestion.id] || currentQuestion.cells;
      return {
        ...prev,
        [currentQuestion.id]: {
          ...base,
          [coord]: {
            ...(base[coord] || {}),
            raw: newRaw,
          },
        },
      };
    });
  };

  const handleSubmitAndGrade = () => {
    if (!currentQuestion) return;
    const answerRaw = currentCells[currentQuestion.targetCell]?.raw ?? '';
    const result = gradeFormulaAnswer(answerRaw, currentQuestion, currentCells);
    setGradesById((prev) => ({
      ...prev,
      [currentQuestion.id]: result,
    }));
  };

  const handleResetSection = () => {
    setGradesById((prev) => {
      const next = { ...prev };
      for (const q of sectionQuestions) {
        delete next[q.id];
      }
      return next;
    });
    setCellsByQuestionId((prev) => {
      const next = { ...prev };
      for (const q of sectionQuestions) {
        delete next[q.id];
      }
      return next;
    });
    setQuestionIndex(0);
    setElapsedSeconds(0);
    setShowHint(false);
    setShowSolution(false);
  };

  const totalSectionPoints = sectionQuestions.reduce((sum, q) => sum + q.points, 0);
  const earnedSectionPoints = sectionQuestions.reduce(
    (sum, q) => sum + (gradesById[q.id]?.scoreAwarded || 0),
    0
  );
  const passedCount = sectionQuestions.filter((q) => gradesById[q.id]?.passed).length;
  const progressPct = totalSectionPoints > 0 ? Math.round((earnedSectionPoints / totalSectionPoints) * 100) : 0;

  const overallPoints = PRACTICAL_TESTS.reduce((sum, q) => sum + q.points, 0);
  const overallEarned = PRACTICAL_TESTS.reduce(
    (sum, q) => sum + (gradesById[q.id]?.scoreAwarded || 0),
    0
  );

  const minutes = String(Math.floor(elapsedSeconds / 60)).padStart(2, '0');
  const seconds = String(elapsedSeconds % 60).padStart(2, '0');
  const activeGrade = currentQuestion ? gradesById[currentQuestion.id] : undefined;

  if (!currentQuestion) return null;

  return (
    <div className="space-y-4">
      {/* Exam Header, Score Tracker & Progress Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          {/* Exam Selector Tabs */}
          <div className="flex flex-wrap items-center gap-1.5" role="tablist" aria-label="Section Practical Exams">
            {SKILL_ORDER.map((lvl) => {
              const t = SKILL_THEMES[lvl];
              const isSelected = lvl === activeSkillLevel;
              return (
                <button
                  key={lvl}
                  type="button"
                  role="tab"
                  aria-selected={isSelected}
                  onClick={() => onSkillLevelChange(lvl)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap ${
                    isSelected
                      ? 'text-white'
                      : 'bg-slate-100 text-slate-700 hover:bg-slate-200/80'
                  }`}
                  style={isSelected ? { backgroundColor: t.hex } : undefined}
                >
                  {t.label} Test
                </button>
              );
            })}
          </div>

          {/* Score & Timer Metrics */}
          <div className="flex flex-wrap items-center gap-4 text-xs">
            <div className="inline-flex items-center gap-1.5 font-mono text-slate-600 tabular-nums">
              <Clock className="w-3.5 h-3.5 text-slate-400" aria-hidden="true" />
              <span>{minutes}:{seconds}</span>
            </div>
            <div className="inline-flex items-center gap-1.5 font-semibold text-slate-900 tabular-nums">
              <Award className="w-4 h-4" style={{ color: theme.hex }} aria-hidden="true" />
              <span>
                Track Score: {earnedSectionPoints} / {totalSectionPoints} pts
              </span>
              <span className="text-slate-400 font-normal">
                ({passedCount}/{sectionQuestions.length} passed)
              </span>
            </div>
            <div className="text-slate-500 tabular-nums">
              Academy Total: <strong className="text-slate-900">{overallEarned}/{overallPoints}</strong>
            </div>
            <button
              type="button"
              onClick={handleResetSection}
              className="inline-flex items-center gap-1 text-slate-600 hover:text-slate-900 font-medium px-2 py-1 rounded hover:bg-slate-100 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" aria-hidden="true" />
              Reset Test
            </button>
          </div>
        </div>

        {/* Visual Progress Bar */}
        <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
          <div
            className="h-full transition-all duration-300"
            style={{ width: `${progressPct}%`, backgroundColor: theme.hex }}
          />
        </div>
      </div>

      {/* Main Assessment Split */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-stretch">
        {/* Left Question & Grading Panel */}
        <section className="lg:col-span-5 bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between gap-4">
          <div className="space-y-4">
            {/* Question Navigator */}
            <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-3">
              <div className="flex items-center gap-1.5">
                {sectionQuestions.map((q, idx) => {
                  const g = gradesById[q.id];
                  const isCurr = idx === questionIndex;
                  return (
                    <button
                      key={q.id}
                      type="button"
                      onClick={() => setQuestionIndex(idx)}
                      className={`px-2.5 py-1 rounded-md text-xs font-semibold border transition-colors cursor-pointer tabular-nums ${
                        isCurr
                          ? 'text-white border-transparent'
                          : g?.passed
                          ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                          : g && !g.passed
                          ? 'bg-red-50 text-red-800 border-red-200'
                          : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                      }`}
                      style={isCurr ? { backgroundColor: theme.hex } : undefined}
                    >
                      Q{idx + 1}
                    </button>
                  );
                })}
              </div>
              <span className="text-xs font-mono font-semibold text-slate-600 tabular-nums">
                {currentQuestion.points} pts · Target {currentQuestion.targetCell}
              </span>
            </div>

            {/* Requirement */}
            <div>
              <h3 className="text-base font-bold text-slate-900">{currentQuestion.title}</h3>
              <p className="mt-1.5 text-sm text-slate-700 leading-relaxed">
                {currentQuestion.requirement}
              </p>
            </div>

            {/* Direct Target Cell Formula Box */}
            <div className="p-3 rounded-lg border border-amber-200 bg-amber-50/70 space-y-2">
              <label
                htmlFor="exam-target-formula"
                className="block text-xs font-bold text-amber-900"
              >
                Your Formula for Cell {currentQuestion.targetCell}
              </label>
              <input
                id="exam-target-formula"
                type="text"
                value={currentCells[currentQuestion.targetCell]?.raw ?? ''}
                onChange={(e) => handleCellChange(currentQuestion.targetCell, e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleSubmitAndGrade();
                  }
                }}
                placeholder={`Type formula starting with = (or use grid fx bar)`}
                className="w-full h-10 px-3 font-mono text-xs bg-white border border-amber-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
              <div className="flex flex-wrap items-center gap-2 pt-1">
                <button
                  type="button"
                  onClick={handleSubmitAndGrade}
                  className="px-4 py-2 rounded-lg text-xs font-semibold text-white transition-opacity hover:opacity-95 cursor-pointer"
                  style={{ backgroundColor: theme.hex }}
                >
                  Submit &amp; Grade
                </button>
                <button
                  type="button"
                  onClick={() => setShowHint((v) => !v)}
                  className="inline-flex items-center gap-1 px-3 py-2 rounded-lg text-xs font-medium border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 cursor-pointer"
                >
                  <HelpCircle className="w-3.5 h-3.5" aria-hidden="true" />
                  {showHint ? 'Hide Hint' : 'Hint'}
                </button>
                <button
                  type="button"
                  onClick={() => setShowSolution((v) => !v)}
                  className="inline-flex items-center gap-1 px-3 py-2 rounded-lg text-xs font-medium border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 cursor-pointer"
                >
                  <Eye className="w-3.5 h-3.5" aria-hidden="true" />
                  {showSolution ? 'Hide Solution' : 'Reveal Solution'}
                </button>
              </div>
            </div>

            {showHint && (
              <div className="p-3 rounded-lg border border-slate-200 bg-slate-50 text-xs text-slate-700">
                <span className="font-semibold text-slate-900">Hint:</span> {currentQuestion.hint}
              </div>
            )}

            {/* Grading Result Badge (Green Pass / Red Fix) */}
            {activeGrade && (
              <div
                role="status"
                className={`p-3.5 rounded-xl border space-y-2 ${
                  activeGrade.passed
                    ? 'border-emerald-200 bg-emerald-50/80 text-emerald-950'
                    : 'border-red-200 bg-red-50/80 text-red-950'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <div className="inline-flex items-center gap-1.5 text-xs font-bold">
                    {activeGrade.passed ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600" aria-hidden="true" />
                    ) : (
                      <XCircle className="w-4 h-4 text-red-600" aria-hidden="true" />
                    )}
                    <span>
                      {activeGrade.passed ? 'PASS' : 'FIX NEEDED'} — {activeGrade.feedbackTitle}
                    </span>
                  </div>
                  <span className="font-mono text-xs font-bold tabular-nums">
                    +{activeGrade.scoreAwarded}/{activeGrade.maxPoints} pts
                  </span>
                </div>

                {activeGrade.feedbackDetails.map((line, i) => (
                  <p key={i} className="text-xs leading-relaxed">
                    {line}
                  </p>
                ))}

                {activeGrade.bestPracticeWarnings.length > 0 && (
                  <ul className="list-disc pl-4 space-y-1 text-xs font-medium">
                    {activeGrade.bestPracticeWarnings.map((w, i) => (
                      <li key={i}>{w}</li>
                    ))}
                  </ul>
                )}
              </div>
            )}

            {/* Optimal Mentor Solution Reveal */}
            {showSolution && (
              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-900 text-slate-100 space-y-2">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs font-semibold text-slate-300">Optimal Mentor Solution</span>
                  <button
                    type="button"
                    onClick={() => handleCellChange(currentQuestion.targetCell, currentQuestion.optimalFormula)}
                    className="text-[11px] font-semibold px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-white border border-slate-700 cursor-pointer"
                  >
                    Insert into {currentQuestion.targetCell}
                  </button>
                </div>
                <div className="font-mono text-xs text-emerald-300 break-all">
                  {currentQuestion.optimalFormula}
                </div>
                <p className="text-xs text-slate-300 leading-relaxed">
                  {currentQuestion.mentorExplanation}
                </p>
              </div>
            )}
          </div>

          {/* Next / Prev Question Footer */}
          <div className="flex items-center justify-between pt-3 border-t border-slate-100">
            <button
              type="button"
              disabled={questionIndex === 0}
              onClick={() => setQuestionIndex((i) => Math.max(0, i - 1))}
              className="px-3 py-1.5 rounded-lg text-xs font-semibold border border-slate-200 text-slate-700 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
            >
              Previous
            </button>
            <span className="text-xs text-slate-500 tabular-nums">
              Question {questionIndex + 1} of {sectionQuestions.length}
            </span>
            <button
              type="button"
              disabled={questionIndex >= sectionQuestions.length - 1}
              onClick={() => setQuestionIndex((i) => Math.min(sectionQuestions.length - 1, i + 1))}
              className="px-3 py-1.5 rounded-lg text-xs font-semibold text-white disabled:opacity-40 cursor-pointer"
              style={{ backgroundColor: theme.hex }}
            >
              Next Question
            </button>
          </div>
        </section>

        {/* Right Interactive Exam Grid */}
        <section className="lg:col-span-7 min-h-[520px]">
          <SpreadsheetGrid
            cells={currentCells}
            onCellChange={handleCellChange}
            selectedCell={selectedCell}
            onSelectCell={setSelectedCell}
            skillLevel={activeSkillLevel}
            yellowTargetCells={[currentQuestion.targetCell]}
            hideAutomationTab
          />
        </section>
      </div>
    </div>
  );
};
