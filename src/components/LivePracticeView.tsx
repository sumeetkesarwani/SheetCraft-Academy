import React, { useMemo, useState } from 'react';
import {
  AlertTriangle,
  ArrowRight,
  BookOpen,
  CornerDownLeft,
  Loader2,
  RotateCcw,
  Terminal,
} from 'lucide-react';
import type {
  AutomationSnippet,
  CellData,
  CoachCardData,
  SheetPreset,
  SkillLevel,
} from '../types/excel';
import { SKILL_THEMES } from '../data/skillThemes';
import { AUTOMATION_LIBRARY } from '../data/practiceScenarios';
import { SpreadsheetGrid } from './SpreadsheetGrid';

interface LivePracticeViewProps {
  skillLevel: SkillLevel;
  preset: SheetPreset;
  cells: Record<string, CellData>;
  selectedCell: string;
  onSelectCell: (coord: string) => void;
  onCellChange: (coord: string, newRaw: string) => void;
  onResetPreset: () => void;
  coachCard: CoachCardData;
  onCoachCardUpdate: (card: CoachCardData) => void;
}

const ARG_COLORS = [
  'text-emerald-700 bg-emerald-50 border-emerald-200',
  'text-blue-700 bg-blue-50 border-blue-200',
  'text-purple-700 bg-purple-50 border-purple-200',
  'text-amber-700 bg-amber-50 border-amber-200',
];

function buildSheetSummary(cells: Record<string, CellData>): string {
  const entries = Object.entries(cells)
    .filter(([, v]) => v.raw.trim() !== '')
    .slice(0, 32)
    .map(([k, v]) => `${k}: ${v.raw}`);
  return entries.join(' | ');
}

export const LivePracticeView: React.FC<LivePracticeViewProps> = ({
  skillLevel,
  preset,
  cells,
  selectedCell,
  onSelectCell,
  onCellChange,
  onResetPreset,
  coachCard,
  onCoachCardUpdate,
}) => {
  const theme = SKILL_THEMES[skillLevel];
  const [question, setQuestion] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [gridTab, setGridTab] = useState<'grid' | 'automation'>('grid');

  const snippets: AutomationSnippet[] = useMemo(() => {
    if (preset.automationSnippets && preset.automationSnippets.length > 0) {
      return preset.automationSnippets;
    }
    return AUTOMATION_LIBRARY;
  }, [preset.automationSnippets]);

  const askCoach = async (promptText: string) => {
    const trimmed = promptText.trim();
    if (!trimmed) return;

    setIsLoading(true);
    setErrorMsg(null);

    try {
      const response = await fetch('/api/coach', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          question: trimmed,
          skillLevel: theme.label,
          moduleTitle: preset.title,
          activeCell: selectedCell,
          sheetSummary: buildSheetSummary(cells),
        }),
      });

      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.coachCard) {
        throw new Error(data.error || 'Unable to reach Live Coach & Lab Desk.');
      }

      const incoming = data.coachCard;
      onCoachCardUpdate({
        id: `cc-${Date.now()}`,
        title: incoming.title || 'Coach Lab Note',
        steps: Array.isArray(incoming.steps) ? incoming.steps : [],
        syntax: incoming.syntax || '',
        syntaxArgs: Array.isArray(incoming.syntaxArgs) ? incoming.syntaxArgs : [],
        pitfallWarning: incoming.pitfallWarning || '',
        nextStepChallenge: incoming.nextStepChallenge || '',
        suggestedFormula: incoming.suggestedFormula || '',
        targetCell: incoming.targetCell || selectedCell,
      });
      setQuestion('');
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : 'Unable to reach Live Coach & Lab Desk.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    askCoach(question);
  };

  const handleApplySuggested = () => {
    if (!coachCard.suggestedFormula) return;
    const dest = coachCard.targetCell || preset.targetCell || selectedCell;
    onSelectCell(dest);
    onCellChange(dest, coachCard.suggestedFormula);
    setGridTab('grid');
  };

  const handleTryNextStep = () => {
    if (coachCard.nextStepChallenge) {
      askCoach(`Walk me through the next step: ${coachCard.nextStepChallenge}`);
    }
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-stretch">
      {/* Left Panel (38% ~ 5/12 cols on desktop): Live Coach & Lab Desk */}
      <section
        aria-label="Live Coach and Lab Notes"
        className="lg:col-span-5 xl:col-span-5 flex flex-col bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs"
      >
        {/* Header with Active Module Badge */}
        <div className="px-4 py-3.5 border-b border-slate-200 bg-slate-50/70">
          <div className="flex items-center justify-between gap-3">
            <div className="flex items-center gap-2 min-w-0">
              <span
                className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-semibold text-white whitespace-nowrap shrink-0"
                style={{ backgroundColor: theme.hex }}
              >
                <BookOpen className="w-3.5 h-3.5" aria-hidden="true" />
                {theme.label}
              </span>
              <span className="text-xs font-medium text-slate-500 truncate">
                Senior FP&amp;A &amp; MIS Desk
              </span>
            </div>
            <button
              type="button"
              onClick={onResetPreset}
              title="Reset active worksheet to default state"
              className="inline-flex items-center gap-1 text-xs font-medium text-slate-600 hover:text-slate-900 px-2 py-1 rounded-md hover:bg-slate-200/60 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" aria-hidden="true" />
              Reset Sheet
            </button>
          </div>
          <h2 className="mt-2 text-base font-bold text-slate-900 leading-snug">
            {preset.title}
          </h2>
          <p className="mt-0.5 text-xs text-slate-600">{preset.description}</p>
        </div>

        {/* 3 Clickable Workplace Scenario Chips */}
        <div className="px-4 py-2.5 border-b border-slate-200 bg-white">
          <div className="text-[11px] font-semibold text-slate-500 mb-1.5">
            Workplace Scenarios
          </div>
          <div className="flex flex-wrap gap-1.5">
            {preset.scenarioChips.map((chip) => (
              <button
                key={chip}
                type="button"
                disabled={isLoading}
                onClick={() => askCoach(chip)}
                className="text-left px-2.5 py-1 rounded-lg text-xs font-medium border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100 hover:border-slate-300 transition-colors cursor-pointer disabled:opacity-50"
              >
                {chip}
              </button>
            ))}
          </div>
        </div>

        {/* Coach Card Content Area */}
        <div className="flex-1 p-3.5 overflow-y-auto excel-scroll space-y-3">
          {errorMsg && (
            <div
              role="alert"
              className="p-3 rounded-lg border border-red-200 bg-red-50 text-xs text-red-800 flex items-start gap-2"
            >
              <AlertTriangle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" aria-hidden="true" />
              <div>
                <div className="font-semibold">Lab Desk Unavailable</div>
                <p className="mt-0.5">{errorMsg}</p>
              </div>
            </div>
          )}

          {isLoading ? (
            <div className="p-5 rounded-xl border border-slate-200 bg-slate-50 space-y-3" aria-busy="true">
              <div className="flex items-center gap-2 text-xs font-semibold text-slate-700">
                <Loader2 className="w-4 h-4 animate-spin" style={{ color: theme.hex }} aria-hidden="true" />
                Reviewing cell {selectedCell} and preparing coaching note...
              </div>
              <div className="h-3 bg-slate-200 rounded w-3/4 animate-pulse" />
              <div className="h-3 bg-slate-200 rounded w-5/6 animate-pulse" />
              <div className="h-9 bg-slate-200 rounded w-full animate-pulse" />
            </div>
          ) : (
            <div className="space-y-4">
              {/* 1. Direct Step-by-Step Logic */}
              <div>
                <h3 className="text-sm font-bold text-slate-900">{coachCard.title}</h3>
                <ol className="mt-2.5 space-y-2">
                  {coachCard.steps.map((step, i) => (
                    <li key={i} className="flex items-start gap-2.5 text-xs leading-relaxed text-slate-700">
                      <span
                        className="w-5 h-5 rounded-md text-[11px] font-bold text-white flex items-center justify-center shrink-0 mt-0.5 tabular-nums"
                        style={{ backgroundColor: theme.hex }}
                      >
                        {i + 1}
                      </span>
                      <span>{step}</span>
                    </li>
                  ))}
                </ol>
              </div>

              {/* 2. Color-Coded Syntax Pill & Arguments */}
              {coachCard.syntax && (
                <div className="p-3 rounded-lg border border-slate-200 bg-slate-50/80 space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold text-slate-500">
                      <Terminal className="w-3.5 h-3.5" aria-hidden="true" />
                      Formula Syntax
                    </span>
                    {coachCard.suggestedFormula && (
                      <button
                        type="button"
                        onClick={handleApplySuggested}
                        className="text-[11px] font-semibold px-2 py-0.5 rounded border border-slate-300 bg-white text-slate-800 hover:bg-slate-100 transition-colors cursor-pointer"
                      >
                        Apply to {coachCard.targetCell || selectedCell}
                      </button>
                    )}
                  </div>
                  <div
                    className="px-3 py-2 rounded-md font-mono text-xs font-medium break-all border"
                    style={{
                      backgroundColor: theme.softHex,
                      borderColor: theme.borderHex,
                      color: '#0F172A',
                    }}
                  >
                    {coachCard.syntax}
                  </div>
                  {coachCard.syntaxArgs.length > 0 && (
                    <div className="flex flex-wrap gap-1.5">
                      {coachCard.syntaxArgs.map((arg, idx) => (
                        <div
                          key={arg.label}
                          className={`px-2 py-1 rounded border text-[11px] leading-snug ${
                            ARG_COLORS[idx % ARG_COLORS.length]
                          }`}
                        >
                          <span className="font-mono font-semibold">{arg.label}</span>
                          <span className="text-slate-600"> — {arg.role}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* 3. Workplace Pitfall Warning Box */}
              {coachCard.pitfallWarning && (
                <div className="p-3 rounded-lg border border-amber-200 bg-amber-50/70">
                  <div className="flex items-center gap-1.5 text-xs font-bold text-amber-900">
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-700" aria-hidden="true" />
                    Workplace Pitfall Warning
                  </div>
                  <p className="mt-1 text-xs leading-relaxed text-amber-900/90">
                    {coachCard.pitfallWarning}
                  </p>
                </div>
              )}

              {/* 4. One-Click "Try Next Step" Challenge Button */}
              {coachCard.nextStepChallenge && (
                <div className="pt-1 flex flex-col gap-2">
                  <div className="text-xs text-slate-600">
                    <span className="font-semibold text-slate-900">Active Drill:</span>{' '}
                    {coachCard.nextStepChallenge}
                  </div>
                  <button
                    type="button"
                    onClick={handleTryNextStep}
                    disabled={isLoading}
                    className="inline-flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-xs font-semibold text-white transition-opacity hover:opacity-95 cursor-pointer disabled:opacity-50"
                    style={{ backgroundColor: theme.hex }}
                  >
                    Try Next Step
                    <ArrowRight className="w-3.5 h-3.5" aria-hidden="true" />
                  </button>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Clean Input Bar */}
        <form onSubmit={handleSubmit} className="p-3 bg-slate-50 border-t border-slate-200">
          <div className="flex items-center gap-2">
            <label htmlFor="coach-question-input" className="sr-only">
              Ask the Live Coach &amp; Lab Desk
            </label>
            <input
              id="coach-question-input"
              type="text"
              value={question}
              onChange={(e) => setQuestion(e.target.value)}
              placeholder="Ask how to write a formula, debug an error, or automate a task..."
              className={`flex-1 h-10 px-3 text-xs bg-white border border-slate-200 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 ${theme.ringClass}`}
            />
            <button
              type="submit"
              disabled={isLoading || !question.trim()}
              className="h-10 px-3.5 rounded-lg text-xs font-semibold text-white inline-flex items-center gap-1.5 transition-opacity hover:opacity-95 disabled:opacity-40 cursor-pointer whitespace-nowrap shrink-0"
              style={{ backgroundColor: theme.hex }}
            >
              <span>Ask Desk</span>
              <CornerDownLeft className="w-3.5 h-3.5" aria-hidden="true" />
            </button>
          </div>
        </form>
      </section>

      {/* Right Panel (62% ~ 7/12 cols on desktop): Interactive Excel Grid & Script Viewer */}
      <section aria-label="Interactive Spreadsheet Sandbox" className="lg:col-span-7 xl:col-span-7 min-h-[540px]">
        <SpreadsheetGrid
          cells={cells}
          onCellChange={onCellChange}
          selectedCell={selectedCell}
          onSelectCell={onSelectCell}
          skillLevel={skillLevel}
          targetCell={preset.targetCell}
          automationSnippets={snippets}
          callouts={preset.callouts}
          activeTabOverride={gridTab}
          onTabChange={setGridTab}
        />
      </section>
    </div>
  );
};
