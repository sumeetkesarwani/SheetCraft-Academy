import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Check, Code2, Copy, Grid3X3, Lightbulb } from 'lucide-react';
import type {
  AutomationSnippet,
  CellData,
  ComputedCell,
  SheetCallout,
  SkillLevel,
} from '../types/excel';
import { SKILL_THEMES } from '../data/skillThemes';
import { evaluateSheet, extractReferencedCells } from '../utils/formulaEngine';

const COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
const ROWS = Array.from({ length: 16 }, (_, i) => i + 1);

interface SpreadsheetGridProps {
  cells: Record<string, CellData>;
  onCellChange: (coord: string, newRaw: string) => void;
  selectedCell: string;
  onSelectCell: (coord: string) => void;
  skillLevel: SkillLevel;
  targetCell?: string;
  yellowTargetCells?: string[];
  automationSnippets?: AutomationSnippet[];
  callouts?: SheetCallout[];
  activeTabOverride?: 'grid' | 'automation';
  onTabChange?: (tab: 'grid' | 'automation') => void;
  hideAutomationTab?: boolean;
}

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  cells,
  onCellChange,
  selectedCell,
  onSelectCell,
  skillLevel,
  targetCell,
  yellowTargetCells = [],
  automationSnippets = [],
  callouts = [],
  activeTabOverride,
  onTabChange,
  hideAutomationTab = false,
}) => {
  const theme = SKILL_THEMES[skillLevel];
  const [internalTab, setInternalTab] = useState<'grid' | 'automation'>('grid');
  const activeTab = activeTabOverride ?? internalTab;

  const setTab = (tab: 'grid' | 'automation') => {
    setInternalTab(tab);
    onTabChange?.(tab);
  };

  const [editingCell, setEditingCell] = useState<string | null>(null);
  const [editDraft, setEditDraft] = useState('');
  const [formulaBarDraft, setFormulaBarDraft] = useState('');
  const [selectedSnippetIdx, setSelectedSnippetIdx] = useState(0);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const inlineInputRef = useRef<HTMLInputElement | null>(null);

  const computedMap: Record<string, ComputedCell> = useMemo(() => evaluateSheet(cells), [cells]);

  const selectedRaw = cells[selectedCell]?.raw ?? '';
  const referencedCells = useMemo(() => {
    const refs = extractReferencedCells(selectedRaw);
    return new Set(refs);
  }, [selectedRaw]);

  useEffect(() => {
    setFormulaBarDraft(cells[selectedCell]?.raw ?? '');
    setEditingCell(null);
  }, [selectedCell, cells]);

  useEffect(() => {
    if (editingCell && inlineInputRef.current) {
      inlineInputRef.current.focus();
      inlineInputRef.current.select();
    }
  }, [editingCell]);

  useEffect(() => {
    setSelectedSnippetIdx(0);
  }, [automationSnippets]);

  const commitCellEdit = (coord: string, value: string) => {
    onCellChange(coord, value);
    setEditingCell(null);
  };

  const handleFormulaBarKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitCellEdit(selectedCell, formulaBarDraft);
    } else if (e.key === 'Escape') {
      setFormulaBarDraft(selectedRaw);
    }
  };

  const handleInlineKeyDown = (e: React.KeyboardEvent<HTMLInputElement>, coord: string) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitCellEdit(coord, editDraft);
      const colIdx = COLS.indexOf(coord[0]);
      const rowNum = parseInt(coord.slice(1), 10);
      if (rowNum < 16 && colIdx !== -1) {
        onSelectCell(`${COLS[colIdx]}${rowNum + 1}`);
      }
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitCellEdit(coord, editDraft);
      const colIdx = COLS.indexOf(coord[0]);
      const rowNum = parseInt(coord.slice(1), 10);
      if (colIdx < COLS.length - 1) {
        onSelectCell(`${COLS[colIdx + 1]}${rowNum}`);
      }
    } else if (e.key === 'Escape') {
      setEditingCell(null);
      setEditDraft(cells[coord]?.raw ?? '');
    }
  };

  const handleCopyCode = async (snippet: AutomationSnippet) => {
    try {
      await navigator.clipboard.writeText(snippet.code);
      setCopiedId(snippet.id);
      setTimeout(() => setCopiedId(null), 1800);
    } catch {
      setCopiedId(null);
    }
  };

  const calloutByCell = useMemo(() => {
    const map = new Map<string, SheetCallout>();
    for (const c of callouts) {
      map.set(c.cell, c);
    }
    return map;
  }, [callouts]);

  const activeSnippet = automationSnippets[selectedSnippetIdx] || automationSnippets[0];

  return (
    <div className="flex flex-col bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs h-full">
      {/* Grid / Automation Code Secondary Tabs */}
      <div className="flex flex-wrap items-center justify-between gap-3 px-3 py-2 bg-slate-50 border-b border-slate-200">
        <div className="flex items-center gap-1 p-0.5 bg-slate-200/70 rounded-lg" role="tablist" aria-label="Workspace View">
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'grid'}
            onClick={() => setTab('grid')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeTab === 'grid'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Grid3X3 className="w-3.5 h-3.5" aria-hidden="true" />
            Spreadsheet Grid
          </button>
          {!hideAutomationTab && (
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === 'automation'}
              onClick={() => setTab('automation')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'automation'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Code2 className="w-3.5 h-3.5" aria-hidden="true" />
              Automation Code
            </button>
          )}
        </div>

        {/* Legend for Cell Highlights */}
        {activeTab === 'grid' && (
          <div className="hidden sm:flex items-center gap-4 text-[11px] text-slate-500">
            {yellowTargetCells.length > 0 ? (
              <span className="inline-flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 rounded-xs bg-amber-100 border border-amber-500" aria-hidden="true" />
                Target Answer Cell
              </span>
            ) : targetCell ? (
              <span className="inline-flex items-center gap-1.5">
                <span
                  className="w-2.5 h-2.5 rounded-xs border"
                  style={{ backgroundColor: theme.softHex, borderColor: theme.hex }}
                  aria-hidden="true"
                />
                Target Formula Cell
              </span>
            ) : null}
            <span className="inline-flex items-center gap-1.5">
              <span className="w-2.5 h-2.5 rounded-xs bg-sky-100 border border-sky-400" aria-hidden="true" />
              Referenced Input
            </span>
          </div>
        )}
      </div>

      {activeTab === 'grid' ? (
        <>
          {/* Authentic Excel Name Box & fx Formula Bar */}
          <div className="flex items-center gap-2 px-3 py-2 bg-white border-b border-slate-200">
            <div
              className="w-14 h-8 flex items-center justify-center rounded-md border border-slate-200 bg-slate-50 font-mono text-xs font-semibold text-slate-800 shrink-0"
              aria-label={`Selected cell ${selectedCell}`}
              data-testid="name-box"
            >
              {selectedCell}
            </div>
            <div className="flex items-center justify-center w-7 h-8 font-mono text-xs italic font-semibold text-slate-500 shrink-0 select-none">
              fx
            </div>
            <input
              type="text"
              aria-label="Formula bar"
              data-testid="formula-bar"
              value={formulaBarDraft}
              onChange={(e) => setFormulaBarDraft(e.target.value)}
              onKeyDown={handleFormulaBarKeyDown}
              placeholder="Enter value or formula starting with = (e.g. =SUM(B2:B4))"
              className={`flex-1 h-8 px-2.5 text-xs font-mono text-slate-900 bg-slate-50/70 border border-slate-200 rounded-md focus:outline-none focus:bg-white focus:ring-2 ${theme.ringClass}`}
            />
            <button
              type="button"
              onClick={() => commitCellEdit(selectedCell, formulaBarDraft)}
              className="h-8 px-3 text-xs font-semibold rounded-md text-white transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              style={{ backgroundColor: theme.hex }}
            >
              Apply
            </button>
          </div>

          {/* Interactive 8x16 Spreadsheet Grid */}
          <div className="flex-1 overflow-auto excel-scroll bg-slate-100/40">
            <table className="w-full border-collapse table-fixed min-w-[740px]" role="grid" aria-label="Interactive Excel Sheet">
              <colgroup>
                <col className="w-10" />
                <col className="w-[140px]" />
                {COLS.slice(1).map((col) => (
                  <col key={col} />
                ))}
              </colgroup>
              <thead className="sticky top-0 z-10">
                <tr>
                  <th className="w-10 h-7 bg-slate-100 border-b border-r border-slate-300 text-[11px] font-medium text-slate-400 select-none" />
                  {COLS.map((col) => {
                    const isActiveCol = selectedCell.startsWith(col);
                    return (
                      <th
                        key={col}
                        scope="col"
                        className={`h-7 border-b border-r border-slate-300 text-[11px] font-semibold select-none text-center ${
                          isActiveCol ? 'bg-slate-200 text-slate-900' : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        {col}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {ROWS.map((row) => {
                  const isActiveRow = selectedCell.slice(1) === String(row);
                  return (
                    <tr key={row}>
                      <th
                        scope="row"
                        className={`w-10 h-8 border-b border-r border-slate-300 text-[11px] font-medium tabular-nums select-none text-center ${
                          isActiveRow ? 'bg-slate-200 text-slate-900 font-semibold' : 'bg-slate-100 text-slate-500'
                        }`}
                      >
                        {row}
                      </th>
                      {COLS.map((col) => {
                        const coord = `${col}${row}`;
                        const rawCell = cells[coord];
                        const comp = computedMap[coord];
                        const isSelected = selectedCell === coord;
                        const isEditing = editingCell === coord;
                        const isYellowTarget = yellowTargetCells.includes(coord);
                        const isThemeTarget = targetCell === coord && !isYellowTarget;
                        const isReferenced = referencedCells.has(coord);
                        const isHeader = rawCell?.format === 'header';
                        const hasCallout = calloutByCell.get(coord);

                        const displayStr = comp ? comp.display : '';
                        const isNumeric =
                          comp &&
                          typeof comp.value === 'number' &&
                          !comp.error &&
                          !isHeader;

                        let cellBg = 'bg-white';
                        if (isHeader) cellBg = 'bg-slate-100/90 font-semibold text-slate-800';
                        if (isYellowTarget) cellBg = 'bg-amber-100/90';
                        else if (isThemeTarget) cellBg = theme.cellTintBg;
                        else if (isReferenced) cellBg = 'bg-sky-50/90';

                        let ringStyle = '';
                        if (isSelected) {
                          ringStyle = isYellowTarget
                            ? 'ring-2 ring-amber-500 z-10'
                            : theme.cellTintBorder + ' z-10';
                        } else if (isYellowTarget) {
                          ringStyle = 'ring-1 ring-amber-400';
                        } else if (isReferenced) {
                          ringStyle = 'ring-1 ring-sky-400/70';
                        }

                        return (
                          <td
                            key={coord}
                            data-testid={`cell-${coord}`}
                            onClick={() => onSelectCell(coord)}
                            onDoubleClick={() => {
                              onSelectCell(coord);
                              setEditingCell(coord);
                              setEditDraft(rawCell?.raw ?? '');
                            }}
                            className={`relative h-8 px-2 border-b border-r border-slate-200 text-xs transition-colors cursor-cell select-none ${cellBg} ${ringStyle} ${
                              comp?.error ? 'text-red-600 font-semibold' : 'text-slate-900'
                            }`}
                          >
                            {isEditing ? (
                              <input
                                ref={inlineInputRef}
                                type="text"
                                aria-label={`Edit cell ${coord}`}
                                value={editDraft}
                                onChange={(e) => setEditDraft(e.target.value)}
                                onBlur={() => commitCellEdit(coord, editDraft)}
                                onKeyDown={(e) => handleInlineKeyDown(e, coord)}
                                className="w-full h-full bg-white font-mono text-xs text-slate-900 focus:outline-none"
                              />
                            ) : (
                              <div
                                className={`flex items-center h-full w-full overflow-hidden ${
                                  isNumeric ? 'justify-end font-mono tabular-nums' : 'justify-start'
                                } ${comp?.isSpilled ? 'italic text-slate-700' : ''}`}
                                title={
                                  comp?.isFormula
                                    ? `${comp.raw} → ${displayStr}`
                                    : comp?.isSpilled
                                    ? `Spilled from ${comp.spillSource}`
                                    : displayStr
                                }
                              >
                                <span className="truncate">{displayStr}</span>
                              </div>
                            )}

                            {/* Visual Callout Indicator for Live Examples */}
                            {hasCallout && (
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  onSelectCell(coord);
                                }}
                                aria-label={`Formula callout ${hasCallout.title} at ${coord}`}
                                className="absolute -top-1 -right-1 w-4 h-4 rounded-full text-[9px] font-bold text-white flex items-center justify-center shadow-xs cursor-pointer"
                                style={{ backgroundColor: theme.hex }}
                                title={`${hasCallout.title}: ${hasCallout.explanation}`}
                              >
                                !
                              </button>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Active Callout Strip when Live Examples callouts exist */}
          {callouts.length > 0 && (
            <div className="px-3 py-2.5 bg-slate-50 border-t border-slate-200">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-[11px] font-semibold text-slate-500 mr-1">Formula Callouts:</span>
                {callouts.map((c) => {
                  const isActive = selectedCell === c.cell;
                  return (
                    <button
                      key={c.id}
                      type="button"
                      onClick={() => onSelectCell(c.cell)}
                      className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium border transition-colors cursor-pointer whitespace-nowrap ${
                        isActive
                          ? 'text-white border-transparent'
                          : 'bg-white text-slate-700 border-slate-200 hover:border-slate-300'
                      }`}
                      style={isActive ? { backgroundColor: theme.hex } : undefined}
                    >
                      <span className="font-mono font-semibold">{c.cell}</span>
                      <span>{c.title}</span>
                    </button>
                  );
                })}
              </div>
              {calloutByCell.get(selectedCell) && (
                <p className="mt-1.5 text-xs text-slate-700">
                  <span className="font-semibold text-slate-900">{selectedCell}:</span>{' '}
                  {calloutByCell.get(selectedCell)?.explanation}
                </p>
              )}
            </div>
          )}
        </>
      ) : (
        /* Automation Code Viewer */
        <div className="flex-1 flex flex-col bg-slate-950 text-slate-100 min-h-[420px]">
          {automationSnippets.length > 0 && activeSnippet ? (
            <>
              <div className="flex flex-wrap items-center justify-between gap-2 px-4 py-3 bg-slate-900 border-b border-slate-800">
                <div className="flex flex-wrap items-center gap-1.5">
                  {automationSnippets.map((snip, idx) => (
                    <button
                      key={snip.id}
                      type="button"
                      onClick={() => setSelectedSnippetIdx(idx)}
                      className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors whitespace-nowrap cursor-pointer ${
                        idx === selectedSnippetIdx
                          ? 'bg-amber-600 text-white'
                          : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                      }`}
                    >
                      {snip.language}
                    </button>
                  ))}
                </div>
                <button
                  type="button"
                  onClick={() => handleCopyCode(activeSnippet)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-100 border border-slate-700 transition-colors cursor-pointer whitespace-nowrap"
                >
                  {copiedId === activeSnippet.id ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-400" aria-hidden="true" />
                      Copied
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" aria-hidden="true" />
                      Copy Code
                    </>
                  )}
                </button>
              </div>
              <div className="px-4 py-3 border-b border-slate-800/80 bg-slate-900/50">
                <div className="text-sm font-semibold text-white">{activeSnippet.title}</div>
                <p className="text-xs text-slate-400 mt-0.5">{activeSnippet.description}</p>
              </div>
              <pre className="flex-1 p-4 overflow-auto excel-scroll font-mono text-xs leading-relaxed text-amber-100/95">
                <code>{activeSnippet.code}</code>
              </pre>
              <div className="px-4 py-3 bg-slate-900 border-t border-slate-800 text-xs text-slate-300 flex items-start gap-2">
                <Lightbulb className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" aria-hidden="true" />
                <span>
                  <strong className="text-white">Field Note:</strong> {activeSnippet.workflowTip}
                </span>
              </div>
            </>
          ) : (
            <div className="flex-1 flex flex-col items-center justify-center p-8 text-center text-slate-400">
              <Code2 className="w-8 h-8 mb-3 text-slate-500" aria-hidden="true" />
              <p className="text-sm font-medium text-slate-200">No automation script attached to this sheet</p>
              <p className="text-xs mt-1 max-w-sm">
                Switch to the Expert Automation track or launch a VBA, Power Query, or =PY() example to inspect code snippets.
              </p>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
