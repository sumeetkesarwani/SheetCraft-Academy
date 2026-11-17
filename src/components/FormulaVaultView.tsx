import React, { useMemo, useState } from 'react';
import { ArrowUpRight, Search } from 'lucide-react';
import type { FormulaCategory, FormulaDefinition, SheetPreset, SkillLevel } from '../types/excel';
import { FORMULA_CATEGORIES, FORMULA_VAULT } from '../data/formulaVaultData';
import { SKILL_THEMES } from '../data/skillThemes';

interface FormulaVaultViewProps {
  activeSkillLevel: SkillLevel;
  onLoadInPractice: (preset: SheetPreset, skillLevel: SkillLevel) => void;
}

const SYNTAX_ARG_CLASSES = [
  'text-emerald-700 font-semibold',
  'text-blue-700 font-semibold',
  'text-purple-700 font-semibold',
  'text-amber-700 font-semibold',
];

export const FormulaVaultView: React.FC<FormulaVaultViewProps> = ({
  activeSkillLevel,
  onLoadInPractice,
}) => {
  const activeTheme = SKILL_THEMES[activeSkillLevel];
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<FormulaCategory | 'All'>('All');
  const [filterByTrack, setFilterByTrack] = useState(false);

  const filteredFormulas = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return FORMULA_VAULT.filter((f) => {
      if (filterByTrack && f.skillLevel !== activeSkillLevel) return false;
      if (selectedCategory !== 'All' && f.category !== selectedCategory) return false;
      if (!q) return true;
      return (
        f.name.toLowerCase().includes(q) ||
        f.category.toLowerCase().includes(q) ||
        f.summary.toLowerCase().includes(q) ||
        f.workplaceTip.toLowerCase().includes(q)
      );
    });
  }, [searchQuery, selectedCategory, filterByTrack, activeSkillLevel]);

  return (
    <div className="space-y-5">
      {/* Search & Filter Controls */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="relative flex-1 max-w-xl">
            <Search
              className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none"
              aria-hidden="true"
            />
            <input
              type="search"
              aria-label="Search formulas in the Formula Vault"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search 70 Excel formulas by name, category, or task (e.g. XLOOKUP, XNPV, FILTER)..."
              className={`w-full h-10 pl-9 pr-3 text-sm bg-slate-50 border border-slate-200 rounded-lg text-slate-900 focus:outline-none focus:bg-white focus:ring-2 ${activeTheme.ringClass}`}
            />
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
              <button
                type="button"
                onClick={() => setFilterByTrack(false)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors cursor-pointer whitespace-nowrap ${
                  !filterByTrack ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                All Tracks ({FORMULA_VAULT.length})
              </button>
              <button
                type="button"
                onClick={() => setFilterByTrack(true)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors cursor-pointer whitespace-nowrap ${
                  filterByTrack ? 'text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'
                }`}
                style={filterByTrack ? { backgroundColor: activeTheme.hex } : undefined}
              >
                {activeTheme.shortLabel} Only
              </button>
            </div>
          </div>
        </div>

        {/* Category Filter Tabs */}
        <div className="flex flex-wrap items-center gap-1.5 pt-1">
          <button
            type="button"
            onClick={() => setSelectedCategory('All')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors cursor-pointer whitespace-nowrap ${
              selectedCategory === 'All'
                ? 'text-white border-transparent'
                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
            }`}
            style={selectedCategory === 'All' ? { backgroundColor: activeTheme.hex } : undefined}
          >
            All Categories
          </button>
          {FORMULA_CATEGORIES.map((cat) => {
            const isSelected = selectedCategory === cat;
            return (
              <button
                key={cat}
                type="button"
                onClick={() => setSelectedCategory(cat)}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors cursor-pointer whitespace-nowrap ${
                  isSelected
                    ? 'text-white border-transparent'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
                style={isSelected ? { backgroundColor: activeTheme.hex } : undefined}
              >
                {cat}
              </button>
            );
          })}
        </div>
      </div>

      {/* Results Count */}
      <div className="flex items-center justify-between text-xs text-slate-500 px-1">
        <span>
          Showing <strong className="text-slate-900 tabular-nums">{filteredFormulas.length}</strong> formulas
          {selectedCategory !== 'All' ? ` in ${selectedCategory}` : ''}
        </span>
      </div>

      {/* Formula Cards Grid */}
      {filteredFormulas.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-xl p-10 text-center">
          <p className="text-sm font-semibold text-slate-900">No matching formulas found</p>
          <p className="text-xs text-slate-500 mt-1">
            Clear your search term or switch back to All Tracks and All Categories.
          </p>
          <button
            type="button"
            onClick={() => {
              setSearchQuery('');
              setSelectedCategory('All');
              setFilterByTrack(false);
            }}
            className="mt-4 px-4 py-2 rounded-lg text-xs font-semibold text-white cursor-pointer"
            style={{ backgroundColor: activeTheme.hex }}
          >
            Reset Filters
          </button>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
          {filteredFormulas.map((formula: FormulaDefinition) => {
            const fTheme = SKILL_THEMES[formula.skillLevel];
            return (
              <article
                key={formula.name}
                className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between gap-3.5 hover:border-slate-300 transition-colors"
              >
                <div className="space-y-3">
                  {/* Quiet Unboxed Metadata Line (Zero-Pill Discipline) */}
                  <div className="flex items-center justify-between gap-2 text-xs">
                    <div className="flex items-center gap-1.5 text-slate-500 truncate">
                      <span className="font-semibold" style={{ color: fTheme.hex }}>
                        {formula.category}
                      </span>
                      <span aria-hidden="true">·</span>
                      <span>{formula.difficulty}</span>
                    </div>
                    <span className="text-[11px] font-medium text-slate-500 whitespace-nowrap shrink-0">
                      {fTheme.shortLabel}
                    </span>
                  </div>

                  {/* Formula Title & Summary */}
                  <div>
                    <h3 className="text-lg font-bold font-mono text-slate-900 tracking-tight">
                      {formula.name}
                    </h3>
                    <p className="mt-1 text-xs leading-relaxed text-slate-600">{formula.summary}</p>
                  </div>

                  {/* Color-Highlighted Syntax Block */}
                  <div
                    className="p-2.5 rounded-lg border font-mono text-xs overflow-x-auto excel-scroll"
                    style={{ backgroundColor: fTheme.softHex, borderColor: fTheme.borderHex }}
                  >
                    <span className="text-slate-900 font-semibold">{formula.syntaxPrefix}</span>
                    {formula.syntaxArgs.map((arg, idx) => (
                      <React.Fragment key={arg.label}>
                        {idx > 0 && <span className="text-slate-500">, </span>}
                        <span className={SYNTAX_ARG_CLASSES[idx % SYNTAX_ARG_CLASSES.length]}>
                          {arg.label}
                        </span>
                      </React.Fragment>
                    ))}
                    <span className="text-slate-900 font-semibold">)</span>
                  </div>

                  {/* Mini Visual Input -> Output Preview Table */}
                  <div className="border border-slate-200 rounded-lg overflow-hidden">
                    <table className="w-full text-[11px] border-collapse">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                          {formula.preview.headers.map((h) => (
                            <th key={h} className="px-2 py-1 text-left font-semibold border-r border-slate-200 last:border-r-0">
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {formula.preview.rows.slice(0, 3).map((row, rIdx) => (
                          <tr key={rIdx} className="border-b border-slate-100 last:border-b-0">
                            {row.map((val, cIdx) => (
                              <td
                                key={cIdx}
                                className={`px-2 py-1 border-r border-slate-100 last:border-r-0 text-slate-700 ${
                                  typeof val === 'number' ? 'font-mono tabular-nums text-right' : ''
                                }`}
                              >
                                {val}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div className="px-2.5 py-1.5 bg-slate-50 border-t border-slate-200 flex items-center justify-between gap-2 text-[11px]">
                      <span className="font-mono text-slate-600 truncate" title={formula.preview.formulaText}>
                        {formula.preview.formulaText}
                      </span>
                      <span className="font-mono font-bold text-slate-900 whitespace-nowrap shrink-0 tabular-nums">
                        → {formula.preview.outputValue}
                      </span>
                    </div>
                  </div>
                </div>

                {/* Action Button */}
                <button
                  type="button"
                  onClick={() => onLoadInPractice(formula.preset, formula.skillLevel)}
                  className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold text-white transition-opacity hover:opacity-95 cursor-pointer"
                  style={{ backgroundColor: fTheme.hex }}
                >
                  Load in Live Practice
                  <ArrowUpRight className="w-3.5 h-3.5" aria-hidden="true" />
                </button>
              </article>
            );
          })}
        </div>
      )}
    </div>
  );
};
