export type SkillLevel =
  | 'foundation'
  | 'data-analyst'
  | 'financial-modeling'
  | 'expert-automation';

export type WorkspaceTab =
  | 'live-practice'
  | 'formula-vault'
  | 'live-examples'
  | 'practical-test';

export type FormulaCategory =
  | 'Math & Aggregation'
  | 'Logical & Error Handling'
  | 'Lookup & Reference'
  | 'Dynamic Arrays'
  | 'Text & Cleaning'
  | 'Date & Time'
  | 'Financial & Modeling';

export type DifficultyTag = 'Core' | 'Intermediate' | 'Advanced' | 'Expert';

export type CellFormat = 'general' | 'currency' | 'percent' | 'number' | 'date' | 'header';

export interface CellData {
  raw: string;
  format?: CellFormat;
  currencySymbol?: string;
  readOnly?: boolean;
  note?: string;
}

export interface ComputedCell {
  raw: string;
  value: string | number | boolean | null;
  display: string;
  isFormula: boolean;
  error?: string;
  format?: CellFormat;
  currencySymbol?: string;
  isSpilled?: boolean;
  spillSource?: string;
}

export interface SyntaxArg {
  label: string;
  role: string;
  optional?: boolean;
}

export interface CoachCardData {
  id: string;
  title: string;
  steps: string[];
  syntax: string;
  syntaxArgs: SyntaxArg[];
  pitfallWarning: string;
  nextStepChallenge: string;
  suggestedFormula?: string;
  targetCell?: string;
}

export type AutomationLanguage = 'Power Query M' | 'VBA Macro' | 'Office Script (TS)' | 'Python (=PY)';

export interface AutomationSnippet {
  id: string;
  title: string;
  language: AutomationLanguage;
  description: string;
  code: string;
  workflowTip: string;
}

export interface SheetCallout {
  id: string;
  cell: string;
  title: string;
  explanation: string;
}

export interface SheetPreset {
  id: string;
  title: string;
  skillLevel: SkillLevel;
  description: string;
  targetCell: string;
  cells: Record<string, CellData>;
  coachCard: CoachCardData;
  scenarioChips: string[];
  automationSnippets?: AutomationSnippet[];
  callouts?: SheetCallout[];
}

export interface FormulaDefinition {
  name: string;
  category: FormulaCategory;
  skillLevel: SkillLevel;
  difficulty: DifficultyTag;
  summary: string;
  syntaxPrefix: string;
  syntaxArgs: SyntaxArg[];
  workplaceTip: string;
  preview: {
    headers: string[];
    rows: (string | number)[][];
    formulaCell: string;
    formulaText: string;
    outputValue: string;
  };
  preset: SheetPreset;
}

export interface LiveExampleCase {
  id: string;
  title: string;
  skillLevel: SkillLevel;
  industry: string;
  keyFunctions: string[];
  summary: string;
  businessOutcome: string;
  preset: SheetPreset;
}

export interface TestQuestion {
  id: string;
  skillLevel: SkillLevel;
  title: string;
  requirement: string;
  targetCell: string;
  points: number;
  cells: Record<string, CellData>;
  expectedValue: string | number;
  numericTolerance?: number;
  requiredPatterns?: { regex: string; message: string }[];
  forbiddenHardcodedNumbers?: number[];
  optimalFormula: string;
  mentorExplanation: string;
  hint: string;
}

export interface GradeResult {
  passed: boolean;
  scoreAwarded: number;
  maxPoints: number;
  computedDisplay: string;
  feedbackTitle: string;
  feedbackDetails: string[];
  bestPracticeWarnings: string[];
}
