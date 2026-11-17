import { useState } from 'react';
import type {
  CellData,
  CoachCardData,
  SheetPreset,
  SkillLevel,
  WorkspaceTab,
} from './types/excel';
import { SKILL_ORDER, SKILL_THEMES } from './data/skillThemes';
import { DEFAULT_PRACTICE_PRESETS } from './data/practiceScenarios';
import { LivePracticeView } from './components/LivePracticeView';
import { FormulaVaultView } from './components/FormulaVaultView';
import { LiveExamplesView } from './components/LiveExamplesView';
import { PracticalTestView } from './components/PracticalTestView';

const WORKSPACE_TABS: { id: WorkspaceTab; label: string }[] = [
  { id: 'live-practice', label: 'Live Practice' },
  { id: 'formula-vault', label: 'Formula Vault' },
  { id: 'live-examples', label: 'Live Examples' },
  { id: 'practical-test', label: 'Practical Test' },
];

export default function App() {
  const [activeTab, setActiveTab] = useState<WorkspaceTab>('live-practice');
  const [skillLevel, setSkillLevel] = useState<SkillLevel>('foundation');

  const [activePreset, setActivePreset] = useState<SheetPreset>(
    DEFAULT_PRACTICE_PRESETS.foundation
  );
  const [practiceCells, setPracticeCells] = useState<Record<string, CellData>>(
    () => structuredClone(DEFAULT_PRACTICE_PRESETS.foundation.cells)
  );
  const [selectedCell, setSelectedCell] = useState<string>(
    DEFAULT_PRACTICE_PRESETS.foundation.targetCell
  );
  const [coachCard, setCoachCard] = useState<CoachCardData>(
    DEFAULT_PRACTICE_PRESETS.foundation.coachCard
  );

  const activeTheme = SKILL_THEMES[skillLevel];

  const handleSelectSkillLevel = (level: SkillLevel) => {
    setSkillLevel(level);
    const nextPreset = DEFAULT_PRACTICE_PRESETS[level];
    setActivePreset(nextPreset);
    setPracticeCells(structuredClone(nextPreset.cells));
    setSelectedCell(nextPreset.targetCell);
    setCoachCard(nextPreset.coachCard);
  };

  const handleLoadPresetInPractice = (preset: SheetPreset, level: SkillLevel) => {
    setSkillLevel(level);
    setActivePreset(preset);
    setPracticeCells(structuredClone(preset.cells));
    setSelectedCell(preset.targetCell || 'E2');
    setCoachCard(preset.coachCard);
    setActiveTab('live-practice');
  };

  const handleResetPracticePreset = () => {
    setPracticeCells(structuredClone(activePreset.cells));
    setSelectedCell(activePreset.targetCell || 'E2');
    setCoachCard(activePreset.coachCard);
  };

  const handleCellChange = (coord: string, newRaw: string) => {
    setPracticeCells((prev) => ({
      ...prev,
      [coord]: {
        ...(prev[coord] || {}),
        raw: newRaw,
      },
    }));
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Strict 3-Zone Top Bar Contract */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-[1440px] mx-auto px-4 sm:px-6 h-14 flex items-center justify-between gap-8">
          {/* Zone 1: Single Text Element Brand Wordmark */}
          <a
            href="#top"
            onClick={(e) => {
              e.preventDefault();
              setActiveTab('live-practice');
            }}
            className="font-display text-lg font-extrabold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
          >
            SheetCraft Academy
          </a>

          {/* Zone 2: 4 Single-Line Workspace Navigation Tabs */}
          <nav
            aria-label="Primary Workspace Navigation"
            className="hidden md:flex items-center gap-6"
          >
            {WORKSPACE_TABS.map((tab) => {
              const isActive = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setActiveTab(tab.id)}
                  className={`relative py-4 text-sm font-medium transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    isActive
                      ? 'text-slate-900 font-semibold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {tab.label}
                  {isActive && (
                    <span
                      className="absolute inset-x-0 bottom-0 h-0.5 rounded-full"
                      style={{ backgroundColor: activeTheme.hex }}
                      aria-hidden="true"
                    />
                  )}
                </button>
              );
            })}
          </nav>

          {/* Zone 3: 1 Primary Action Button */}
          <div className="flex items-center shrink-0">
            <button
              type="button"
              onClick={() =>
                setActiveTab(activeTab === 'practical-test' ? 'live-practice' : 'practical-test')
              }
              className="px-4 py-2 text-xs font-semibold text-white rounded-lg transition-opacity hover:opacity-95 whitespace-nowrap shrink-0 cursor-pointer"
              style={{ backgroundColor: activeTheme.hex }}
            >
              {activeTab === 'practical-test' ? 'Return to Lab' : 'Take Practical Test'}
            </button>
          </div>
        </div>

        {/* Mobile Compact Tab Strip */}
        <nav
          aria-label="Mobile Workspace Navigation"
          className="md:hidden border-t border-slate-100 px-2 py-1.5 flex items-center justify-between gap-1"
        >
          {WORKSPACE_TABS.map((tab) => {
            const isActive = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`flex-1 text-center px-1.5 py-1.5 rounded-md text-[11px] font-semibold whitespace-nowrap shrink-0 cursor-pointer ${
                  isActive ? 'text-white' : 'text-slate-600 hover:bg-slate-100'
                }`}
                style={isActive ? { backgroundColor: activeTheme.hex } : undefined}
              >
                {tab.label}
              </button>
            );
          })}
        </nav>
      </header>

      {/* Main Content Container */}
      <main id="top" className="flex-1 max-w-[1440px] w-full mx-auto px-4 sm:px-6 py-5 space-y-5">
        {/* Skill Level Filter Strip (4 Color-Coded Tracks) */}
        <section
          aria-label="Skill Level Tracks"
          className="bg-white border border-slate-200 rounded-xl p-3 sm:px-4 sm:py-3 flex flex-col lg:flex-row lg:items-center justify-between gap-3"
        >
          <div className="flex flex-wrap items-center gap-2" role="group" aria-label="Filter by Skill Level">
            {SKILL_ORDER.map((lvl) => {
              const t = SKILL_THEMES[lvl];
              const isSelected = skillLevel === lvl;
              return (
                <button
                  key={lvl}
                  type="button"
                  aria-pressed={isSelected}
                  onClick={() => handleSelectSkillLevel(lvl)}
                  className={`px-3.5 py-2 rounded-lg text-xs font-semibold border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    isSelected
                      ? 'text-white border-transparent shadow-xs'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                  }`}
                  style={
                    isSelected
                      ? { backgroundColor: t.hex }
                      : { borderColor: `${t.hex}33` }
                  }
                >
                  {t.label}
                </button>
              );
            })}
          </div>

          <div className="flex items-center gap-2 text-xs text-slate-600 min-w-0">
            <span
              className="font-semibold whitespace-nowrap shrink-0"
              style={{ color: activeTheme.hex }}
            >
              {activeTheme.tagline}:
            </span>
            <span className="truncate">{activeTheme.topics}</span>
          </div>
        </section>

        {/* Active Workspace View */}
        {activeTab === 'live-practice' && (
          <LivePracticeView
            skillLevel={skillLevel}
            preset={activePreset}
            cells={practiceCells}
            selectedCell={selectedCell}
            onSelectCell={setSelectedCell}
            onCellChange={handleCellChange}
            onResetPreset={handleResetPracticePreset}
            coachCard={coachCard}
            onCoachCardUpdate={setCoachCard}
          />
        )}

        {activeTab === 'formula-vault' && (
          <FormulaVaultView
            activeSkillLevel={skillLevel}
            onLoadInPractice={handleLoadPresetInPractice}
          />
        )}

        {activeTab === 'live-examples' && (
          <LiveExamplesView
            activeSkillLevel={skillLevel}
            onSkillLevelChange={handleSelectSkillLevel}
            onOpenInLivePractice={handleLoadPresetInPractice}
          />
        )}

        {activeTab === 'practical-test' && (
          <PracticalTestView
            activeSkillLevel={skillLevel}
            onSkillLevelChange={handleSelectSkillLevel}
          />
        )}
      </main>

      {/* Clean Quiet Footer */}
      <footer className="border-t border-slate-200 bg-white mt-8">
        <div className="max-w-[1440px] mx-auto px-4 sm:px-6 py-4 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-slate-500">
          <span>SheetCraft Academy · Foundation to Expert Automation</span>
          <span>Live Coach &amp; Lab Desk · Formula Vault · Practical Exams</span>
        </div>
      </footer>
    </div>
  );
}
