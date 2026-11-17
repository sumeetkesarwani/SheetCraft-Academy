import React, { useMemo, useState } from 'react';
import { ArrowLeft, ArrowUpRight, Layers, Play } from 'lucide-react';
import type { CellData, LiveExampleCase, SheetPreset, SkillLevel } from '../types/excel';
import { LIVE_EXAMPLES } from '../data/liveExamplesData';
import { SKILL_ORDER, SKILL_THEMES } from '../data/skillThemes';
import { SpreadsheetGrid } from './SpreadsheetGrid';

interface LiveExamplesViewProps {
  activeSkillLevel: SkillLevel;
  onSkillLevelChange: (level: SkillLevel) => void;
  onOpenInLivePractice: (preset: SheetPreset, skillLevel: SkillLevel) => void;
}

export const LiveExamplesView: React.FC<LiveExamplesViewProps> = ({
  activeSkillLevel,
  onSkillLevelChange,
  onOpenInLivePractice,
}) => {
  const [launchedExample, setLaunchedExample] = useState<LiveExampleCase | null>(null);
  const [exampleCells, setExampleCells] = useState<Record<string, CellData>>({});
  const [selectedCell, setSelectedCell] = useState<string>('E2');
  const [showAllTracks, setShowAllTracks] = useState(true);

  const visibleExamples = useMemo(() => {
    if (showAllTracks) return LIVE_EXAMPLES;
    return LIVE_EXAMPLES.filter((ex) => ex.skillLevel === activeSkillLevel);
  }, [showAllTracks, activeSkillLevel]);

  const handleLaunch = (ex: LiveExampleCase) => {
    onSkillLevelChange(ex.skillLevel);
    setLaunchedExample(ex);
    setExampleCells(structuredClone(ex.preset.cells));
    setSelectedCell(ex.preset.targetCell || 'E2');
  };

  if (launchedExample) {
    const theme = SKILL_THEMES[launchedExample.skillLevel];
    const callouts = launchedExample.preset.callouts || [];

    return (
      <div className="space-y-4">
        {/* Top Context Bar */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="space-y-1">
            <div className="flex items-center gap-2 text-xs text-slate-500">
              <button
                type="button"
                onClick={() => setLaunchedExample(null)}
                className="inline-flex items-center gap-1 font-semibold text-slate-700 hover:text-slate-900 cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" aria-hidden="true" />
                Back to Case Gallery
              </button>
              <span aria-hidden="true">·</span>
              <span className="font-semibold" style={{ color: theme.hex }}>
                {theme.label}
              </span>
              <span aria-hidden="true">·</span>
              <span>{launchedExample.industry}</span>
            </div>
            <h2 className="text-lg font-bold text-slate-900">{launchedExample.title}</h2>
            <p className="text-xs text-slate-600">{launchedExample.summary}</p>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() =>
                onOpenInLivePractice(
                  { ...launchedExample.preset, cells: exampleCells },
                  launchedExample.skillLevel
                )
              }
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-semibold text-white transition-opacity hover:opacity-95 cursor-pointer whitespace-nowrap"
              style={{ backgroundColor: theme.hex }}
            >
              Open in Live Practice Coach
              <ArrowUpRight className="w-3.5 h-3.5" aria-hidden="true" />
            </button>
          </div>
        </div>

        {/* Interactive Split: Callout Walkthrough + Full Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-stretch">
          <aside className="lg:col-span-4 bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between gap-4">
            <div className="space-y-4">
              <div>
                <div className="text-xs font-semibold text-slate-500">Formula Architecture</div>
                <h3 className="text-sm font-bold text-slate-900 mt-0.5">
                  Step-by-Step Cell Callouts
                </h3>
                <p className="text-xs text-slate-600 mt-1">
                  Click any marker below or in the grid to highlight its formula and referenced input cells.
                </p>
              </div>

              <div className="space-y-2.5">
                {callouts.map((c, idx) => {
                  const isActive = selectedCell === c.cell;
                  return (
                    <button
                      key={c.id}
                      type="button"
                      onClick={() => setSelectedCell(c.cell)}
                      className={`w-full text-left p-3 rounded-lg border transition-colors cursor-pointer ${
                        isActive
                          ? 'border-transparent'
                          : 'border-slate-200 bg-slate-50/70 hover:bg-slate-100/80'
                      }`}
                      style={
                        isActive
                          ? { backgroundColor: theme.softHex, borderColor: theme.hex }
                          : undefined
                      }
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-bold text-slate-900">
                          0{idx + 1}. {c.title}
                        </span>
                        <span
                          className="font-mono text-[11px] font-semibold px-2 py-0.5 rounded text-white"
                          style={{ backgroundColor: theme.hex }}
                        >
                          {c.cell}
                        </span>
                      </div>
                      <p className="mt-1 text-xs text-slate-600 leading-relaxed">{c.explanation}</p>
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="p-3 rounded-lg border border-slate-200 bg-slate-50 text-xs text-slate-700">
              <div className="font-semibold text-slate-900">Business Outcome</div>
              <p className="mt-1 leading-relaxed">{launchedExample.businessOutcome}</p>
            </div>
          </aside>

          <div className="lg:col-span-8 min-h-[520px]">
            <SpreadsheetGrid
              cells={exampleCells}
              onCellChange={(coord, raw) =>
                setExampleCells((prev) => ({
                  ...prev,
                  [coord]: { ...(prev[coord] || {}), raw },
                }))
              }
              selectedCell={selectedCell}
              onSelectCell={setSelectedCell}
              skillLevel={launchedExample.skillLevel}
              targetCell={launchedExample.preset.targetCell}
              automationSnippets={launchedExample.preset.automationSnippets}
              callouts={callouts}
            />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-5">
      {/* Gallery Header Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h2 className="text-base font-bold text-slate-900">
            Interactive Real-World Business Cases
          </h2>
          <p className="text-xs text-slate-600 mt-0.5">
            Launch production templates with working datasets, live formulas, and step-by-step cell callouts.
          </p>
        </div>

        <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg self-start sm:self-auto shrink-0">
          <button
            type="button"
            onClick={() => setShowAllTracks(true)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors cursor-pointer whitespace-nowrap ${
              showAllTracks ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            All Sections ({LIVE_EXAMPLES.length})
          </button>
          <button
            type="button"
            onClick={() => setShowAllTracks(false)}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors cursor-pointer whitespace-nowrap ${
              !showAllTracks
                ? 'text-white shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
            style={!showAllTracks ? { backgroundColor: SKILL_THEMES[activeSkillLevel].hex } : undefined}
          >
            {SKILL_THEMES[activeSkillLevel].shortLabel} Only
          </button>
        </div>
      </div>

      {/* Grouped Sections */}
      <div className="space-y-6">
        {SKILL_ORDER.filter((lvl) => showAllTracks || lvl === activeSkillLevel).map((lvl) => {
          const theme = SKILL_THEMES[lvl];
          const groupCases = visibleExamples.filter((ex) => ex.skillLevel === lvl);
          if (groupCases.length === 0) return null;

          return (
            <section key={lvl} className="space-y-3">
              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                <div className="flex items-center gap-2">
                  <Layers className="w-4 h-4" style={{ color: theme.hex }} aria-hidden="true" />
                  <h3 className="text-sm font-bold text-slate-900">{theme.label}</h3>
                  <span className="text-xs text-slate-500">· {theme.tagline}</span>
                </div>
                <span className="text-xs font-mono text-slate-500 tabular-nums">
                  {groupCases.length} templates
                </span>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {groupCases.map((ex) => (
                  <article
                    key={ex.id}
                    className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between gap-4 hover:border-slate-300 transition-colors"
                  >
                    <div className="space-y-2.5">
                      <div className="flex items-center gap-1.5 text-xs text-slate-500">
                        <span className="font-semibold" style={{ color: theme.hex }}>
                          {theme.shortLabel}
                        </span>
                        <span aria-hidden="true">·</span>
                        <span>{ex.industry}</span>
                      </div>

                      <h4 className="text-base font-bold text-slate-900 leading-snug">
                        {ex.title}
                      </h4>
                      <p className="text-xs text-slate-600 leading-relaxed">{ex.summary}</p>

                      <div className="pt-1 text-xs font-mono text-slate-700">
                        {ex.keyFunctions.join(' · ')}
                      </div>
                    </div>

                    <div className="space-y-3 pt-2 border-t border-slate-100">
                      <p className="text-[11px] text-slate-500 leading-relaxed">
                        {ex.businessOutcome}
                      </p>
                      <button
                        type="button"
                        onClick={() => handleLaunch(ex)}
                        className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold text-white transition-opacity hover:opacity-95 cursor-pointer"
                        style={{ backgroundColor: theme.hex }}
                      >
                        <Play className="w-3.5 h-3.5" aria-hidden="true" />
                        Launch Interactive Example
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            </section>
          );
        })}
      </div>
    </div>
  );
};
