import type { SkillLevel } from '../types/excel';

export interface SkillTheme {
  id: SkillLevel;
  label: string;
  shortLabel: string;
  tagline: string;
  topics: string;
  hex: string;
  softHex: string;
  borderHex: string;
  activeBgClass: string;
  activeTextClass: string;
  softBgClass: string;
  softTextClass: string;
  borderClass: string;
  ringClass: string;
  cellTintBg: string;
  cellTintBorder: string;
  progressBarClass: string;
}

export const SKILL_THEMES: Record<SkillLevel, SkillTheme> = {
  foundation: {
    id: 'foundation',
    label: 'Foundation',
    shortLabel: 'Foundation',
    tagline: 'Core Mechanics & Clean Sheets',
    topics: 'Cell locking ($), basic math, IF/IFS, text cleaning, dates, conditional rules',
    hex: '#107C41',
    softHex: '#ECFDF5',
    borderHex: '#A7F3D0',
    activeBgClass: 'bg-[#107C41] text-white',
    activeTextClass: 'text-[#107C41]',
    softBgClass: 'bg-[#ECFDF5]',
    softTextClass: 'text-[#107C41]',
    borderClass: 'border-[#107C41]/25',
    ringClass: 'focus-visible:ring-[#107C41]',
    cellTintBg: 'bg-emerald-50/90',
    cellTintBorder: 'ring-2 ring-[#107C41]',
    progressBarClass: 'bg-[#107C41]',
  },
  'data-analyst': {
    id: 'data-analyst',
    label: 'Data Analyst',
    shortLabel: 'Data Analyst',
    tagline: 'Lookups & Dynamic Arrays',
    topics: 'SUMIFS/COUNTIFS, XLOOKUP, INDEX/MATCH, FILTER, UNIQUE, SORT, PIVOTBY',
    hex: '#1D4ED8',
    softHex: '#EFF6FF',
    borderHex: '#BFDBFE',
    activeBgClass: 'bg-[#1D4ED8] text-white',
    activeTextClass: 'text-[#1D4ED8]',
    softBgClass: 'bg-[#EFF6FF]',
    softTextClass: 'text-[#1D4ED8]',
    borderClass: 'border-[#1D4ED8]/25',
    ringClass: 'focus-visible:ring-[#1D4ED8]',
    cellTintBg: 'bg-blue-50/90',
    cellTintBorder: 'ring-2 ring-[#1D4ED8]',
    progressBarClass: 'bg-[#1D4ED8]',
  },
  'financial-modeling': {
    id: 'financial-modeling',
    label: 'Financial Modeling',
    shortLabel: 'Financial Modeling',
    tagline: 'Valuation, Cash Flow & Bridges',
    topics: 'LET, LAMBDA, XNPV/XIRR, 3-statement links, variance bridges, sensitivity tables',
    hex: '#6D28D9',
    softHex: '#F5F3FF',
    borderHex: '#DDD6FE',
    activeBgClass: 'bg-[#6D28D9] text-white',
    activeTextClass: 'text-[#6D28D9]',
    softBgClass: 'bg-[#F5F3FF]',
    softTextClass: 'text-[#6D28D9]',
    borderClass: 'border-[#6D28D9]/25',
    ringClass: 'focus-visible:ring-[#6D28D9]',
    cellTintBg: 'bg-purple-50/90',
    cellTintBorder: 'ring-2 ring-[#6D28D9]',
    progressBarClass: 'bg-[#6D28D9]',
  },
  'expert-automation': {
    id: 'expert-automation',
    label: 'Expert Automation',
    shortLabel: 'Automation',
    tagline: 'Power Query, VBA, Scripts & Python',
    topics: 'Power Query (M), VBA Macros, TypeScript Office Scripts, Python (=PY())',
    hex: '#B45309',
    softHex: '#FFFBEB',
    borderHex: '#FDE68A',
    activeBgClass: 'bg-[#B45309] text-white',
    activeTextClass: 'text-[#B45309]',
    softBgClass: 'bg-[#FFFBEB]',
    softTextClass: 'text-[#B45309]',
    borderClass: 'border-[#B45309]/25',
    ringClass: 'focus-visible:ring-[#B45309]',
    cellTintBg: 'bg-amber-50/90',
    cellTintBorder: 'ring-2 ring-[#B45309]',
    progressBarClass: 'bg-[#B45309]',
  },
};

export const SKILL_ORDER: SkillLevel[] = [
  'foundation',
  'data-analyst',
  'financial-modeling',
  'expert-automation',
];
