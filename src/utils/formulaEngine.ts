import type { CellData, CellFormat, ComputedCell, GradeResult, TestQuestion } from '../types/excel';

export type ScalarValue = string | number | boolean | null;
export type MatrixValue = ScalarValue[][];
export type EvalValue = ScalarValue | MatrixValue;

export class ExcelError extends Error {
  code: string;
  constructor(code: string) {
    super(code);
    this.code = code;
  }
}

export function isMatrix(val: unknown): val is MatrixValue {
  return Array.isArray(val);
}

export function toScalar(val: EvalValue): ScalarValue {
  if (isMatrix(val)) {
    return val[0]?.[0] ?? null;
  }
  return val;
}

export function flattenValues(val: EvalValue): ScalarValue[] {
  if (!isMatrix(val)) return [val];
  const out: ScalarValue[] = [];
  for (const row of val) {
    for (const cell of row) {
      out.push(cell);
    }
  }
  return out;
}

export function colLetterToIndex(col: string): number {
  let idx = 0;
  const upper = col.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    idx = idx * 26 + (upper.charCodeAt(i) - 64);
  }
  return idx - 1;
}

export function indexToColLetter(index: number): string {
  let n = index + 1;
  let col = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

export function parseCellCoord(ref: string): { col: number; row: number; coord: string } | null {
  const match = ref.trim().match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!match) return null;
  const colStr = match[1].toUpperCase();
  const rowNum = parseInt(match[2], 10);
  if (rowNum < 1) return null;
  return {
    col: colLetterToIndex(colStr),
    row: rowNum,
    coord: `${colStr}${rowNum}`,
  };
}

export function expandRange(startRef: string, endRef: string): { coords: string[][]; flat: string[] } | null {
  const start = parseCellCoord(startRef);
  const end = parseCellCoord(endRef);
  if (!start || !end) return null;
  const minCol = Math.min(start.col, end.col);
  const maxCol = Math.max(start.col, end.col);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  const coords: string[][] = [];
  const flat: string[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    const rowCoords: string[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      const coord = `${indexToColLetter(c)}${r}`;
      rowCoords.push(coord);
      flat.push(coord);
    }
    coords.push(rowCoords);
  }
  return { coords, flat };
}

export function extractReferencedCells(formula: string): string[] {
  if (!formula || !formula.trim().startsWith('=')) return [];
  // Strip quoted strings first so text inside quotes doesn't match cell refs
  const stripped = formula.replace(/"([^"]|"")*"/g, '""');
  const refs = new Set<string>();

  // Match ranges like $A$2:$D$20 or A2:B5
  const rangeRegex = /\$?[A-Za-z]{1,2}\$?\d+\s*:\s*\$?[A-Za-z]{1,2}\$?\d+/g;
  let withoutRanges = stripped;
  const rangeMatches = stripped.match(rangeRegex) || [];
  for (const rm of rangeMatches) {
    const [a, b] = rm.split(':');
    const expanded = expandRange(a, b);
    if (expanded) {
      for (const c of expanded.flat) {
        refs.add(c);
      }
    }
  }
  withoutRanges = withoutRanges.replace(rangeRegex, ' ');

  // Match single cell references like $E$2, A1, B12
  const singleRegex = /\$?[A-Za-z]{1,2}\$?\d+/g;
  const singleMatches = withoutRanges.match(singleRegex) || [];
  for (const sm of singleMatches) {
    const parsed = parseCellCoord(sm);
    if (parsed) refs.add(parsed.coord);
  }
  return Array.from(refs);
}

// Date utilities (ISO YYYY-MM-DD <-> Excel serial day)
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

function parseDateValue(val: ScalarValue): Date | null {
  if (typeof val === 'number' && val > 20000 && val < 90000) {
    return new Date(EXCEL_EPOCH_UTC + val * 86400000);
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      const [y, m, d] = trimmed.split('-').map(Number);
      return new Date(Date.UTC(y, m - 1, d));
    }
    const parsed = Date.parse(trimmed);
    if (!Number.isNaN(parsed)) {
      const dt = new Date(parsed);
      return new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()));
    }
  }
  return null;
}

function dateToSerial(date: Date): number {
  const utc = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  return Math.round((utc - EXCEL_EPOCH_UTC) / 86400000);
}

function formatDateISO(date: Date): string {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function toNumber(val: ScalarValue): number {
  if (typeof val === 'number') {
    if (Number.isNaN(val)) throw new ExcelError('#VALUE!');
    return val;
  }
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (val === null || val === '') return 0;
  if (typeof val === 'string') {
    const cleaned = val.trim().replace(/[₹$,]/g, '');
    if (cleaned.endsWith('%')) {
      const n = Number(cleaned.slice(0, -1));
      if (!Number.isNaN(n)) return n / 100;
    }
    if (cleaned === '') return 0;
    const n = Number(cleaned);
    if (!Number.isNaN(n)) return n;
    const dt = parseDateValue(val);
    if (dt) return dateToSerial(dt);
  }
  throw new ExcelError('#VALUE!');
}

function toStrictNumber(val: ScalarValue): number | null {
  if (typeof val === 'number' && !Number.isNaN(val)) return val;
  if (typeof val === 'string') {
    const cleaned = val.trim().replace(/[₹$,]/g, '');
    if (!cleaned) return null;
    if (cleaned.endsWith('%')) {
      const n = Number(cleaned.slice(0, -1));
      return Number.isNaN(n) ? null : n / 100;
    }
    const n = Number(cleaned);
    return Number.isNaN(n) ? null : n;
  }
  return null;
}

function toBoolean(val: ScalarValue): boolean {
  if (typeof val === 'boolean') return val;
  if (typeof val === 'number') return val !== 0;
  if (typeof val === 'string') {
    const u = val.trim().toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    if (u === '') return false;
    throw new ExcelError('#VALUE!');
  }
  return false;
}

function toStringVal(val: ScalarValue): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  return String(val);
}

function parseRawScalar(raw: string): ScalarValue {
  const trimmed = raw.trim();
  if (trimmed === '') return null;
  if (trimmed.toUpperCase() === 'TRUE') return true;
  if (trimmed.toUpperCase() === 'FALSE') return false;
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
  if (/^-?\d+(\.\d+)?%$/.test(trimmed)) return Number(trimmed.slice(0, -1)) / 100;
  if (/^[₹$]-?\d[\d,]*(\.\d+)?$/.test(trimmed)) {
    return Number(trimmed.slice(1).replace(/,/g, ''));
  }
  return raw;
}

// Criteria matching for SUMIF / SUMIFS / COUNTIF / COUNTIFS
function matchesCriteria(cellVal: ScalarValue, criteria: ScalarValue): boolean {
  if (criteria === null) return cellVal === null || cellVal === '';
  const critStr = toStringVal(criteria).trim();

  const opMatch = critStr.match(/^(<=|>=|<>|!=|<|>|=)(.*)$/);
  if (opMatch) {
    const op = opMatch[1];
    const rhsRaw = opMatch[2].trim();
    const rhsNum = toStrictNumber(rhsRaw);
    const cellNum = toStrictNumber(cellVal);

    if (rhsNum !== null && cellNum !== null) {
      switch (op) {
        case '=': return cellNum === rhsNum;
        case '<>':
        case '!=': return cellNum !== rhsNum;
        case '<': return cellNum < rhsNum;
        case '>': return cellNum > rhsNum;
        case '<=': return cellNum <= rhsNum;
        case '>=': return cellNum >= rhsNum;
      }
    }
    const lhsStr = toStringVal(cellVal).toLowerCase();
    const rhsStr = rhsRaw.toLowerCase();
    switch (op) {
      case '=': return wildcardMatch(lhsStr, rhsStr);
      case '<>':
      case '!=': return !wildcardMatch(lhsStr, rhsStr);
      case '<': return lhsStr < rhsStr;
      case '>': return lhsStr > rhsStr;
      case '<=': return lhsStr <= rhsStr;
      case '>=': return lhsStr >= rhsStr;
    }
  }

  const critNum = toStrictNumber(criteria);
  const cellNum = toStrictNumber(cellVal);
  if (critNum !== null && cellNum !== null) {
    return Math.abs(cellNum - critNum) < 1e-9;
  }

  return wildcardMatch(toStringVal(cellVal).toLowerCase(), critStr.toLowerCase());
}

function wildcardMatch(text: string, pattern: string): boolean {
  if (!pattern.includes('*') && !pattern.includes('?')) {
    return text === pattern;
  }
  const escaped = pattern
    .replace(/[.+^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*')
    .replace(/\?/g, '.');
  return new RegExp(`^${escaped}$`, 'i').test(text);
}

// Tokenizer & AST Parser
type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'CELL'
  | 'IDENT'
  | 'OP'
  | 'LPAREN'
  | 'RPAREN'
  | 'LBRACE'
  | 'RBRACE'
  | 'COMMA'
  | 'SEMICOLON'
  | 'COLON'
  | 'EOF';

interface Token {
  type: TokenType;
  value: string;
}

function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const len = input.length;

  while (i < len) {
    const ch = input[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // String literal
    if (ch === '"') {
      i++;
      let str = '';
      while (i < len) {
        if (input[i] === '"') {
          if (i + 1 < len && input[i + 1] === '"') {
            str += '"';
            i += 2;
          } else {
            i++;
            break;
          }
        } else {
          str += input[i];
          i++;
        }
      }
      tokens.push({ type: 'STRING', value: str });
      continue;
    }

    // Multi-char operators
    const two = input.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>' || two === '!=') {
      tokens.push({ type: 'OP', value: two });
      i += 2;
      continue;
    }

    if ('+-*/^&=<>'.includes(ch)) {
      tokens.push({ type: 'OP', value: ch });
      i++;
      continue;
    }

    if (ch === '(') { tokens.push({ type: 'LPAREN', value: '(' }); i++; continue; }
    if (ch === ')') { tokens.push({ type: 'RPAREN', value: ')' }); i++; continue; }
    if (ch === '{') { tokens.push({ type: 'LBRACE', value: '{' }); i++; continue; }
    if (ch === '}') { tokens.push({ type: 'RBRACE', value: '}' }); i++; continue; }
    if (ch === ',') { tokens.push({ type: 'COMMA', value: ',' }); i++; continue; }
    if (ch === ';') { tokens.push({ type: 'SEMICOLON', value: ';' }); i++; continue; }
    if (ch === ':') { tokens.push({ type: 'COLON', value: ':' }); i++; continue; }

    // Numbers
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(input[i + 1] || ''))) {
      let numStr = '';
      while (i < len && /[\d.]/.test(input[i])) {
        numStr += input[i++];
      }
      if (i < len && (input[i] === 'e' || input[i] === 'E') && /[+\-\d]/.test(input[i + 1] || '')) {
        numStr += input[i++];
        if (input[i] === '+' || input[i] === '-') numStr += input[i++];
        while (i < len && /\d/.test(input[i])) numStr += input[i++];
      }
      if (i < len && input[i] === '%') {
        i++;
        tokens.push({ type: 'NUMBER', value: String(Number(numStr) / 100) });
      } else {
        tokens.push({ type: 'NUMBER', value: numStr });
      }
      continue;
    }

    // Identifiers or Cell references ($A$1, A1, XLOOKUP, TRUE, FALSE)
    if (/[A-Za-z_$]/.test(ch)) {
      let id = '';
      while (i < len && /[A-Za-z0-9_$.]/.test(input[i])) {
        id += input[i++];
      }
      const upper = id.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({ type: 'BOOLEAN', value: upper });
      } else if (/^\$?[A-Za-z]{1,2}\$?\d+$/.test(id)) {
        tokens.push({ type: 'CELL', value: id });
      } else {
        tokens.push({ type: 'IDENT', value: upper });
      }
      continue;
    }

    throw new ExcelError('#NAME?');
  }

  tokens.push({ type: 'EOF', value: '' });
  return tokens;
}

type ASTNode =
  | { kind: 'literal'; value: ScalarValue }
  | { kind: 'cell'; ref: string }
  | { kind: 'range'; start: string; end: string }
  | { kind: 'var'; name: string }
  | { kind: 'unary'; op: string; expr: ASTNode }
  | { kind: 'binary'; op: string; left: ASTNode; right: ASTNode }
  | { kind: 'call'; name: string; args: ASTNode[] }
  | { kind: 'invoke'; callee: ASTNode; args: ASTNode[] }
  | { kind: 'array'; rows: ASTNode[][] };

class Parser {
  tokens: Token[];
  pos = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  peek(): Token {
    return this.tokens[this.pos] || { type: 'EOF', value: '' };
  }

  consume(): Token {
    return this.tokens[this.pos++];
  }

  parse(): ASTNode {
    const node = this.parseComparison();
    if (this.peek().type !== 'EOF') {
      throw new ExcelError('#ERROR!');
    }
    return node;
  }

  parseComparison(): ASTNode {
    let left = this.parseConcat();
    while (
      this.peek().type === 'OP' &&
      ['=', '<>', '!=', '<', '>', '<=', '>='].includes(this.peek().value)
    ) {
      const op = this.consume().value;
      const right = this.parseConcat();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  parseConcat(): ASTNode {
    let left = this.parseAddSub();
    while (this.peek().type === 'OP' && this.peek().value === '&') {
      const op = this.consume().value;
      const right = this.parseAddSub();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  parseAddSub(): ASTNode {
    let left = this.parseMulDiv();
    while (this.peek().type === 'OP' && (this.peek().value === '+' || this.peek().value === '-')) {
      const op = this.consume().value;
      const right = this.parseMulDiv();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  parseMulDiv(): ASTNode {
    let left = this.parsePower();
    while (this.peek().type === 'OP' && (this.peek().value === '*' || this.peek().value === '/')) {
      const op = this.consume().value;
      const right = this.parsePower();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  parsePower(): ASTNode {
    let left = this.parseUnary();
    while (this.peek().type === 'OP' && this.peek().value === '^') {
      const op = this.consume().value;
      const right = this.parseUnary();
      left = { kind: 'binary', op, left, right };
    }
    return left;
  }

  parseUnary(): ASTNode {
    if (this.peek().type === 'OP' && (this.peek().value === '-' || this.peek().value === '+')) {
      const op = this.consume().value;
      const expr = this.parseUnary();
      return { kind: 'unary', op, expr };
    }
    return this.parsePrimary();
  }

  parsePrimary(): ASTNode {
    const tok = this.peek();

    if (tok.type === 'NUMBER') {
      this.consume();
      return { kind: 'literal', value: Number(tok.value) };
    }
    if (tok.type === 'STRING') {
      this.consume();
      return { kind: 'literal', value: tok.value };
    }
    if (tok.type === 'BOOLEAN') {
      this.consume();
      return { kind: 'literal', value: tok.value === 'TRUE' };
    }
    if (tok.type === 'CELL') {
      this.consume();
      if (this.peek().type === 'COLON') {
        this.consume();
        const endTok = this.peek();
        if (endTok.type !== 'CELL') throw new ExcelError('#REF!');
        this.consume();
        return { kind: 'range', start: tok.value, end: endTok.value };
      }
      return { kind: 'cell', ref: tok.value };
    }
    if (tok.type === 'IDENT') {
      this.consume();
      if (this.peek().type === 'LPAREN') {
        this.consume();
        const args: ASTNode[] = [];
        if (this.peek().type !== 'RPAREN') {
          args.push(this.parseComparison());
          while (this.peek().type === 'COMMA') {
            this.consume();
            args.push(this.parseComparison());
          }
        }
        if (this.peek().type !== 'RPAREN') throw new ExcelError('#ERROR!');
        this.consume();
        let node: ASTNode = { kind: 'call', name: tok.value, args };
        // Support immediate invocation e.g. LAMBDA(x, y, x+y)(3, 4)
        while (this.peek().type === 'LPAREN') {
          this.consume();
          const invokeArgs: ASTNode[] = [];
          if (this.peek().type !== 'RPAREN') {
            invokeArgs.push(this.parseComparison());
            while (this.peek().type === 'COMMA') {
              this.consume();
              invokeArgs.push(this.parseComparison());
            }
          }
          if (this.peek().type !== 'RPAREN') throw new ExcelError('#ERROR!');
          this.consume();
          node = { kind: 'invoke', callee: node, args: invokeArgs };
        }
        return node;
      }
      return { kind: 'var', name: tok.value };
    }
    if (tok.type === 'LPAREN') {
      this.consume();
      const expr = this.parseComparison();
      if (this.peek().type !== 'RPAREN') throw new ExcelError('#ERROR!');
      this.consume();
      return expr;
    }
    if (tok.type === 'LBRACE') {
      this.consume();
      const rows: ASTNode[][] = [];
      let currentRow: ASTNode[] = [];
      if (this.peek().type !== 'RBRACE') {
        currentRow.push(this.parseComparison());
        while (this.peek().type === 'COMMA' || this.peek().type === 'SEMICOLON') {
          const sep = this.consume();
          if (sep.type === 'SEMICOLON') {
            rows.push(currentRow);
            currentRow = [];
          }
          currentRow.push(this.parseComparison());
        }
        rows.push(currentRow);
      }
      if (this.peek().type !== 'RBRACE') throw new ExcelError('#ERROR!');
      this.consume();
      return { kind: 'array', rows };
    }

    throw new ExcelError('#ERROR!');
  }
}

interface LambdaFn {
  __isLambda: true;
  params: string[];
  body: ASTNode;
  scope: Map<string, EvalValue | LambdaFn>;
}

type ScopeValue = EvalValue | LambdaFn;

function isLambda(val: ScopeValue): val is LambdaFn {
  return typeof val === 'object' && val !== null && '__isLambda' in val;
}

// Financial math helpers
function calcNPV(rate: number, cashflows: number[]): number {
  let total = 0;
  for (let i = 0; i < cashflows.length; i++) {
    total += cashflows[i] / Math.pow(1 + rate, i + 1);
  }
  return total;
}

function calcXNPV(rate: number, values: number[], dates: Date[]): number {
  if (values.length !== dates.length || values.length === 0) throw new ExcelError('#VALUE!');
  const d0 = dates[0].getTime();
  let total = 0;
  for (let i = 0; i < values.length; i++) {
    const days = (dates[i].getTime() - d0) / 86400000;
    total += values[i] / Math.pow(1 + rate, days / 365);
  }
  return total;
}

function calcIRR(values: number[], guess = 0.1): number {
  let r = guess;
  for (let iter = 0; iter < 100; iter++) {
    let f = 0;
    let df = 0;
    for (let i = 0; i < values.length; i++) {
      const denom = Math.pow(1 + r, i);
      f += values[i] / denom;
      if (i > 0) {
        df -= (i * values[i]) / Math.pow(1 + r, i + 1);
      }
    }
    if (Math.abs(f) < 1e-7) return r;
    if (Math.abs(df) < 1e-12) break;
    r = r - f / df;
  }
  throw new ExcelError('#NUM!');
}

function calcXIRR(values: number[], dates: Date[], guess = 0.1): number {
  if (values.length !== dates.length || values.length < 2) throw new ExcelError('#VALUE!');
  const d0 = dates[0].getTime();
  let r = guess;
  for (let iter = 0; iter < 100; iter++) {
    let f = 0;
    let df = 0;
    for (let i = 0; i < values.length; i++) {
      const t = (dates[i].getTime() - d0) / (365 * 86400000);
      const factor = Math.pow(1 + r, t);
      f += values[i] / factor;
      if (t !== 0) {
        df -= (t * values[i]) / Math.pow(1 + r, t + 1);
      }
    }
    if (Math.abs(f) < 1e-7) return r;
    if (Math.abs(df) < 1e-12) break;
    r = r - f / df;
  }
  throw new ExcelError('#NUM!');
}

function calcPMT(rate: number, nper: number, pv: number, fv = 0, type = 0): number {
  if (rate === 0) return -(pv + fv) / nper;
  const pvif = Math.pow(1 + rate, nper);
  let pmt = (rate / (pvif - 1)) * -(pv * pvif + fv);
  if (type === 1) pmt /= 1 + rate;
  return pmt;
}

function calcFV(rate: number, nper: number, pmt: number, pv = 0, type = 0): number {
  if (rate === 0) return -pv - pmt * nper;
  const term = Math.pow(1 + rate, nper);
  return -pv * term - pmt * (1 + rate * type) * ((term - 1) / rate);
}

function calcPV(rate: number, nper: number, pmt: number, fv = 0, type = 0): number {
  if (rate === 0) return -fv - pmt * nper;
  const term = Math.pow(1 + rate, nper);
  return (-fv - pmt * (1 + rate * type) * ((term - 1) / rate)) / term;
}

function formatNumberDisplay(num: number, format?: CellFormat, currencySymbol = '₹'): string {
  if (!Number.isFinite(num)) return '#NUM!';
  if (format === 'percent') {
    return `${(num * 100).toFixed(1)}%`;
  }
  if (format === 'currency') {
    const abs = Math.abs(num);
    const formatted = abs.toLocaleString('en-IN', {
      minimumFractionDigits: abs % 1 === 0 ? 0 : 2,
      maximumFractionDigits: 2,
    });
    return num < 0 ? `-${currencySymbol}${formatted}` : `${currencySymbol}${formatted}`;
  }
  if (format === 'number') {
    return num.toLocaleString('en-IN', {
      minimumFractionDigits: num % 1 === 0 ? 0 : 2,
      maximumFractionDigits: 2,
    });
  }
  if (Math.abs(num - Math.round(num)) < 1e-9) {
    return String(Math.round(num));
  }
  return Number(num.toFixed(4)).toString();
}

export function formatCellValue(
  val: ScalarValue,
  format?: CellFormat,
  currencySymbol = '₹'
): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  if (typeof val === 'number') {
    if (format === 'date') {
      const dt = parseDateValue(val);
      if (dt) return formatDateISO(dt);
    }
    return formatNumberDisplay(val, format, currencySymbol);
  }
  return String(val);
}

export function evaluateSheet(cells: Record<string, CellData>): Record<string, ComputedCell> {
  const computed: Record<string, ComputedCell> = {};
  const visiting = new Set<string>();
  const rawValuesCache = new Map<string, EvalValue>();

  function getCellEval(coord: string): EvalValue {
    const normalized = parseCellCoord(coord)?.coord ?? coord.toUpperCase();
    if (rawValuesCache.has(normalized)) {
      return rawValuesCache.get(normalized)!;
    }
    const cell = cells[normalized];
    if (!cell || cell.raw.trim() === '') {
      rawValuesCache.set(normalized, null);
      return null;
    }
    const raw = cell.raw.trim();
    if (!raw.startsWith('=')) {
      const scalar = parseRawScalar(raw);
      rawValuesCache.set(normalized, scalar);
      return scalar;
    }

    if (visiting.has(normalized)) {
      throw new ExcelError('#CIRC!');
    }
    visiting.add(normalized);
    try {
      const exprText = raw.slice(1).trim();
      const tokens = tokenize(exprText);
      const ast = new Parser(tokens).parse();
      const result = evalNode(ast, new Map());
      if (isLambda(result)) {
        throw new ExcelError('#CALC!');
      }
      rawValuesCache.set(normalized, result);
      return result;
    } finally {
      visiting.delete(normalized);
    }
  }

  function evalToScalar(node: ASTNode, scope: Map<string, ScopeValue>): ScalarValue {
    const val = evalNode(node, scope);
    if (isLambda(val)) throw new ExcelError('#VALUE!');
    return toScalar(val);
  }

  function evalRangeMatrix(node: ASTNode, scope: Map<string, ScopeValue>): MatrixValue {
    if (node.kind === 'range') {
      const expanded = expandRange(node.start, node.end);
      if (!expanded) throw new ExcelError('#REF!');
      return expanded.coords.map((rowCoords) =>
        rowCoords.map((c) => toScalar(getCellEval(c)))
      );
    }
    const res = evalNode(node, scope);
    if (isLambda(res)) throw new ExcelError('#VALUE!');
    if (isMatrix(res)) return res;
    return [[res]];
  }

  function evalRangeFlat(node: ASTNode, scope: Map<string, ScopeValue>): ScalarValue[] {
    const matrix = evalRangeMatrix(node, scope);
    return flattenValues(matrix);
  }

  function evalNode(node: ASTNode, scope: Map<string, ScopeValue>): EvalValue | LambdaFn {
    switch (node.kind) {
      case 'literal':
        return node.value;

      case 'cell': {
        const parsed = parseCellCoord(node.ref);
        if (!parsed) throw new ExcelError('#REF!');
        return getCellEval(parsed.coord);
      }

      case 'range': {
        const expanded = expandRange(node.start, node.end);
        if (!expanded) throw new ExcelError('#REF!');
        return expanded.coords.map((rowCoords) =>
          rowCoords.map((c) => toScalar(getCellEval(c)))
        );
      }

      case 'var': {
        if (scope.has(node.name)) {
          return scope.get(node.name)!;
        }
        throw new ExcelError('#NAME?');
      }

      case 'array': {
        return node.rows.map((r) => r.map((item) => evalToScalar(item, scope)));
      }

      case 'unary': {
        const val = toNumber(evalToScalar(node.expr, scope));
        return node.op === '-' ? -val : val;
      }

      case 'binary': {
        const lVal = evalNode(node.left, scope);
        const rVal = evalNode(node.right, scope);
        if (isLambda(lVal) || isLambda(rVal)) throw new ExcelError('#VALUE!');

        const applyScalarBinary = (lRaw: ScalarValue, rRaw: ScalarValue): ScalarValue => {
          if (node.op === '&') {
            return toStringVal(lRaw) + toStringVal(rRaw);
          }
          if (['=', '<>', '!=', '<', '>', '<=', '>='].includes(node.op)) {
            const lNum = toStrictNumber(lRaw);
            const rNum = toStrictNumber(rRaw);
            if (lNum !== null && rNum !== null) {
              switch (node.op) {
                case '=': return Math.abs(lNum - rNum) < 1e-9;
                case '<>':
                case '!=': return Math.abs(lNum - rNum) >= 1e-9;
                case '<': return lNum < rNum;
                case '>': return lNum > rNum;
                case '<=': return lNum <= rNum;
                case '>=': return lNum >= rNum;
              }
            }
            const lStr = toStringVal(lRaw).toLowerCase();
            const rStr = toStringVal(rRaw).toLowerCase();
            switch (node.op) {
              case '=': return lStr === rStr;
              case '<>':
              case '!=': return lStr !== rStr;
              case '<': return lStr < rStr;
              case '>': return lStr > rStr;
              case '<=': return lStr <= rStr;
              case '>=': return lStr >= rStr;
            }
          }
          const l = toNumber(lRaw);
          const r = toNumber(rRaw);
          switch (node.op) {
            case '+': return l + r;
            case '-': return l - r;
            case '*': return l * r;
            case '/':
              if (r === 0) throw new ExcelError('#DIV/0!');
              return l / r;
            case '^': return Math.pow(l, r);
            default:
              throw new ExcelError('#ERROR!');
          }
        };

        if (isMatrix(lVal) || isMatrix(rVal)) {
          const lMat: MatrixValue = isMatrix(lVal) ? lVal : [[lVal]];
          const rMat: MatrixValue = isMatrix(rVal) ? rVal : [[rVal]];
          const rows = Math.max(lMat.length, rMat.length);
          const cols = Math.max(lMat[0]?.length || 1, rMat[0]?.length || 1);
          const out: MatrixValue = [];
          for (let r = 0; r < rows; r++) {
            const row: ScalarValue[] = [];
            for (let c = 0; c < cols; c++) {
              const lv = lMat[lMat.length === 1 ? 0 : r]?.[lMat[0].length === 1 ? 0 : c] ?? null;
              const rv = rMat[rMat.length === 1 ? 0 : r]?.[rMat[0].length === 1 ? 0 : c] ?? null;
              row.push(applyScalarBinary(lv, rv));
            }
            out.push(row);
          }
          return out;
        }

        return applyScalarBinary(lVal, rVal);
      }

      case 'invoke': {
        const fnVal = evalNode(node.callee, scope);
        if (!isLambda(fnVal)) throw new ExcelError('#VALUE!');
        if (node.args.length !== fnVal.params.length) throw new ExcelError('#VALUE!');
        const childScope = new Map(fnVal.scope);
        fnVal.params.forEach((param, idx) => {
          childScope.set(param, evalNode(node.args[idx], scope));
        });
        return evalNode(fnVal.body, childScope);
      }

      case 'call':
        return evalFunction(node.name, node.args, scope);
    }
  }

  function evalFunction(name: string, args: ASTNode[], scope: Map<string, ScopeValue>): EvalValue | LambdaFn {
    const fn = name.toUpperCase();

    // Check if invoking a LET-bound LAMBDA variable
    if (scope.has(fn)) {
      const bound = scope.get(fn)!;
      if (isLambda(bound)) {
        if (args.length !== bound.params.length) throw new ExcelError('#VALUE!');
        const childScope = new Map(bound.scope);
        bound.params.forEach((p, idx) => {
          childScope.set(p, evalNode(args[idx], scope));
        });
        return evalNode(bound.body, childScope);
      }
    }

    const getNumericItems = (argNodes: ASTNode[]): number[] => {
      const nums: number[] = [];
      for (const a of argNodes) {
        if (a.kind === 'range' || a.kind === 'array') {
          for (const v of evalRangeFlat(a, scope)) {
            if (typeof v === 'number' && !Number.isNaN(v)) nums.push(v);
          }
        } else {
          const val = evalNode(a, scope);
          if (isMatrix(val)) {
            for (const v of flattenValues(val)) {
              if (typeof v === 'number' && !Number.isNaN(v)) nums.push(v);
            }
          } else if (!isLambda(val) && val !== null && val !== '') {
            const n = toStrictNumber(val);
            if (n !== null) nums.push(n);
          }
        }
      }
      return nums;
    };

    switch (fn) {
      case 'PY': {
        // Preview simulation for Python in Excel =PY(...) expressions
        return '[Python DataFrame Ready]';
      }

      case 'LET': {
        if (args.length < 3 || args.length % 2 === 0) throw new ExcelError('#VALUE!');
        const childScope = new Map(scope);
        for (let i = 0; i < args.length - 1; i += 2) {
          const nameNode = args[i];
          if (nameNode.kind !== 'var') throw new ExcelError('#VALUE!');
          const val = evalNode(args[i + 1], childScope);
          childScope.set(nameNode.name, val);
        }
        return evalNode(args[args.length - 1], childScope);
      }

      case 'LAMBDA': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const params: string[] = [];
        for (let i = 0; i < args.length - 1; i++) {
          const p = args[i];
          if (p.kind !== 'var') throw new ExcelError('#VALUE!');
          params.push(p.name);
        }
        return {
          __isLambda: true,
          params,
          body: args[args.length - 1],
          scope: new Map(scope),
        };
      }

      case 'SUM': {
        const nums = getNumericItems(args);
        return nums.reduce((acc, n) => acc + n, 0);
      }

      case 'AVERAGE': {
        const nums = getNumericItems(args);
        if (nums.length === 0) throw new ExcelError('#DIV/0!');
        return nums.reduce((a, b) => a + b, 0) / nums.length;
      }

      case 'MIN': {
        const nums = getNumericItems(args);
        return nums.length === 0 ? 0 : Math.min(...nums);
      }

      case 'MAX': {
        const nums = getNumericItems(args);
        return nums.length === 0 ? 0 : Math.max(...nums);
      }

      case 'COUNT': {
        return getNumericItems(args).length;
      }

      case 'COUNTA': {
        let count = 0;
        for (const a of args) {
          for (const v of evalRangeFlat(a, scope)) {
            if (v !== null && v !== '') count++;
          }
        }
        return count;
      }

      case 'SUBTOTAL': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const funcNum = toNumber(evalToScalar(args[0], scope)) % 100;
        const nums = getNumericItems(args.slice(1));
        switch (funcNum) {
          case 1: return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
          case 2:
          case 3: return nums.length;
          case 4: return nums.length ? Math.max(...nums) : 0;
          case 5: return nums.length ? Math.min(...nums) : 0;
          case 9: return nums.reduce((a, b) => a + b, 0);
          default: return nums.reduce((a, b) => a + b, 0);
        }
      }

      case 'AGGREGATE': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const funcNum = toNumber(evalToScalar(args[0], scope));
        const nums = getNumericItems(args.slice(2));
        if (funcNum === 9) return nums.reduce((a, b) => a + b, 0);
        if (funcNum === 1) return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        if (funcNum === 4) return nums.length ? Math.max(...nums) : 0;
        if (funcNum === 5) return nums.length ? Math.min(...nums) : 0;
        return nums.reduce((a, b) => a + b, 0);
      }

      case 'ROUND': {
        const num = toNumber(evalToScalar(args[0], scope));
        const digits = args[1] ? Math.trunc(toNumber(evalToScalar(args[1], scope))) : 0;
        const factor = Math.pow(10, digits);
        return Math.round((num + Number.EPSILON) * factor) / factor;
      }

      case 'ROUNDUP': {
        const num = toNumber(evalToScalar(args[0], scope));
        const digits = args[1] ? Math.trunc(toNumber(evalToScalar(args[1], scope))) : 0;
        const factor = Math.pow(10, digits);
        return num >= 0 ? Math.ceil(num * factor) / factor : Math.floor(num * factor) / factor;
      }

      case 'ROUNDDOWN': {
        const num = toNumber(evalToScalar(args[0], scope));
        const digits = args[1] ? Math.trunc(toNumber(evalToScalar(args[1], scope))) : 0;
        const factor = Math.pow(10, digits);
        return Math.trunc(num * factor) / factor;
      }

      case 'ABS':
        return Math.abs(toNumber(evalToScalar(args[0], scope)));

      case 'MOD': {
        const n = toNumber(evalToScalar(args[0], scope));
        const d = toNumber(evalToScalar(args[1], scope));
        if (d === 0) throw new ExcelError('#DIV/0!');
        return n - d * Math.floor(n / d);
      }

      case 'IF': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const cond = toBoolean(evalToScalar(args[0], scope));
        if (cond) return evalNode(args[1], scope) as EvalValue;
        return args[2] ? (evalNode(args[2], scope) as EvalValue) : false;
      }

      case 'IFS': {
        for (let i = 0; i + 1 < args.length; i += 2) {
          if (toBoolean(evalToScalar(args[i], scope))) {
            return evalNode(args[i + 1], scope) as EvalValue;
          }
        }
        throw new ExcelError('#N/A');
      }

      case 'AND': {
        for (const a of args) {
          for (const v of evalRangeFlat(a, scope)) {
            if (v !== null && !toBoolean(v)) return false;
          }
        }
        return true;
      }

      case 'OR': {
        for (const a of args) {
          for (const v of evalRangeFlat(a, scope)) {
            if (v !== null && toBoolean(v)) return true;
          }
        }
        return false;
      }

      case 'NOT':
        return !toBoolean(evalToScalar(args[0], scope));

      case 'SWITCH': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const expr = evalToScalar(args[0], scope);
        for (let i = 1; i + 1 < args.length; i += 2) {
          const cand = evalToScalar(args[i], scope);
          if (toStringVal(expr).toLowerCase() === toStringVal(cand).toLowerCase()) {
            return evalNode(args[i + 1], scope) as EvalValue;
          }
        }
        if (args.length % 2 === 0) {
          return evalNode(args[args.length - 1], scope) as EvalValue;
        }
        throw new ExcelError('#N/A');
      }

      case 'IFERROR': {
        if (args.length !== 2) throw new ExcelError('#VALUE!');
        try {
          return evalNode(args[0], scope) as EvalValue;
        } catch {
          return evalNode(args[1], scope) as EvalValue;
        }
      }

      case 'IFNA': {
        if (args.length !== 2) throw new ExcelError('#VALUE!');
        try {
          return evalNode(args[0], scope) as EvalValue;
        } catch (err) {
          if (err instanceof ExcelError && err.code === '#N/A') {
            return evalNode(args[1], scope) as EvalValue;
          }
          throw err;
        }
      }

      case 'CHOOSE': {
        const idx = Math.trunc(toNumber(evalToScalar(args[0], scope)));
        if (idx < 1 || idx >= args.length) throw new ExcelError('#VALUE!');
        return evalNode(args[idx], scope) as EvalValue;
      }

      case 'VLOOKUP': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const lookupVal = evalToScalar(args[0], scope);
        const table = evalRangeMatrix(args[1], scope);
        const colIdx = Math.trunc(toNumber(evalToScalar(args[2], scope)));
        const exact = args[3] ? !toBoolean(evalToScalar(args[3], scope)) : false;
        if (colIdx < 1 || colIdx > (table[0]?.length || 0)) throw new ExcelError('#REF!');

        if (exact) {
          for (const row of table) {
            if (matchesCriteria(row[0], lookupVal)) {
              return row[colIdx - 1] ?? null;
            }
          }
          throw new ExcelError('#N/A');
        } else {
          // Approximate match (ascending sorted first column)
          const targetNum = toStrictNumber(lookupVal);
          let bestRow: ScalarValue[] | null = null;
          for (const row of table) {
            const cellNum = toStrictNumber(row[0]);
            if (targetNum !== null && cellNum !== null) {
              if (cellNum <= targetNum) bestRow = row;
            } else if (toStringVal(row[0]).toLowerCase() <= toStringVal(lookupVal).toLowerCase()) {
              bestRow = row;
            }
          }
          if (!bestRow) throw new ExcelError('#N/A');
          return bestRow[colIdx - 1] ?? null;
        }
      }

      case 'HLOOKUP': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const lookupVal = evalToScalar(args[0], scope);
        const table = evalRangeMatrix(args[1], scope);
        const rowIdx = Math.trunc(toNumber(evalToScalar(args[2], scope)));
        const exact = args[3] ? !toBoolean(evalToScalar(args[3], scope)) : false;
        if (rowIdx < 1 || rowIdx > table.length) throw new ExcelError('#REF!');
        const firstRow = table[0] || [];
        if (exact) {
          for (let c = 0; c < firstRow.length; c++) {
            if (matchesCriteria(firstRow[c], lookupVal)) {
              return table[rowIdx - 1]?.[c] ?? null;
            }
          }
          throw new ExcelError('#N/A');
        }
        let bestCol = -1;
        const targetNum = toStrictNumber(lookupVal);
        for (let c = 0; c < firstRow.length; c++) {
          const cellNum = toStrictNumber(firstRow[c]);
          if (targetNum !== null && cellNum !== null && cellNum <= targetNum) {
            bestCol = c;
          }
        }
        if (bestCol === -1) throw new ExcelError('#N/A');
        return table[rowIdx - 1]?.[bestCol] ?? null;
      }

      case 'XLOOKUP': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const lookupVal = evalToScalar(args[0], scope);
        const lookupArr = evalRangeMatrix(args[1], scope);
        const returnArr = evalRangeMatrix(args[2], scope);
        const matchMode = args[4] ? Math.trunc(toNumber(evalToScalar(args[4], scope))) : 0;
        const searchMode = args[5] ? Math.trunc(toNumber(evalToScalar(args[5], scope))) : 1;

        const isRowVector = lookupArr.length === 1 && (lookupArr[0]?.length || 0) > 1;
        const keys: ScalarValue[] = isRowVector
          ? lookupArr[0]
          : lookupArr.map((r) => r[0]);

        let matchIdx = -1;
        const indices = searchMode === -1
          ? Array.from({ length: keys.length }, (_, i) => keys.length - 1 - i)
          : Array.from({ length: keys.length }, (_, i) => i);

        if (matchMode === 0 || matchMode === 2) {
          for (const i of indices) {
            if (matchesCriteria(keys[i], lookupVal)) {
              matchIdx = i;
              break;
            }
          }
        } else if (matchMode === -1) {
          const targetNum = toStrictNumber(lookupVal);
          let bestVal = -Infinity;
          for (const i of indices) {
            const kNum = toStrictNumber(keys[i]);
            if (targetNum !== null && kNum !== null && kNum <= targetNum && kNum >= bestVal) {
              bestVal = kNum;
              matchIdx = i;
            }
          }
        } else if (matchMode === 1) {
          const targetNum = toStrictNumber(lookupVal);
          let bestVal = Infinity;
          for (const i of indices) {
            const kNum = toStrictNumber(keys[i]);
            if (targetNum !== null && kNum !== null && kNum >= targetNum && kNum <= bestVal) {
              bestVal = kNum;
              matchIdx = i;
            }
          }
        }

        if (matchIdx === -1) {
          if (args[3]) return evalNode(args[3], scope) as EvalValue;
          throw new ExcelError('#N/A');
        }

        if (isRowVector) {
          if (returnArr.length === 1) return returnArr[0][matchIdx] ?? null;
          return returnArr.map((r) => [r[matchIdx] ?? null]);
        } else {
          const row = returnArr[matchIdx];
          if (!row) throw new ExcelError('#REF!');
          return row.length === 1 ? (row[0] ?? null) : [row];
        }
      }

      case 'INDEX': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const arr = evalRangeMatrix(args[0], scope);
        const rowNum = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        const colNum = args[2] ? Math.trunc(toNumber(evalToScalar(args[2], scope))) : 1;

        if (rowNum === 0 && colNum > 0) {
          return arr.map((r) => [r[colNum - 1] ?? null]);
        }
        if (colNum === 0 && rowNum > 0) {
          return [arr[rowNum - 1] ?? []];
        }
        if (rowNum < 1 || rowNum > arr.length || colNum < 1 || colNum > (arr[0]?.length || 0)) {
          throw new ExcelError('#REF!');
        }
        return arr[rowNum - 1][colNum - 1] ?? null;
      }

      case 'MATCH':
      case 'XMATCH': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const lookupVal = evalToScalar(args[0], scope);
        const arr = evalRangeFlat(args[1], scope);
        const matchType = args[2] ? Math.trunc(toNumber(evalToScalar(args[2], scope))) : (fn === 'XMATCH' ? 0 : 1);

        if (matchType === 0 || matchType === 2) {
          for (let i = 0; i < arr.length; i++) {
            if (matchesCriteria(arr[i], lookupVal)) return i + 1;
          }
          throw new ExcelError('#N/A');
        }
        const targetNum = toStrictNumber(lookupVal);
        if (matchType === 1 || matchType === -1) {
          for (let i = 0; i < arr.length; i++) {
            if (matchesCriteria(arr[i], lookupVal)) return i + 1;
          }
          let bestIdx = -1;
          for (let i = 0; i < arr.length; i++) {
            const vNum = toStrictNumber(arr[i]);
            if (targetNum !== null && vNum !== null) {
              if (matchType === 1 && vNum <= targetNum) bestIdx = i + 1;
              if (matchType === -1 && vNum >= targetNum) bestIdx = i + 1;
            }
          }
          if (bestIdx !== -1) return bestIdx;
        }
        throw new ExcelError('#N/A');
      }

      case 'SUMIF': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const range = evalRangeFlat(args[0], scope);
        const criteria = evalToScalar(args[1], scope);
        const sumRange = args[2] ? evalRangeFlat(args[2], scope) : range;
        let total = 0;
        for (let i = 0; i < range.length; i++) {
          if (matchesCriteria(range[i], criteria)) {
            const n = toStrictNumber(sumRange[i] ?? null);
            if (n !== null) total += n;
          }
        }
        return total;
      }

      case 'SUMIFS': {
        if (args.length < 3 || args.length % 2 === 0) throw new ExcelError('#VALUE!');
        const sumRange = evalRangeFlat(args[0], scope);
        const pairs: { range: ScalarValue[]; crit: ScalarValue }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          pairs.push({
            range: evalRangeFlat(args[i], scope),
            crit: evalToScalar(args[i + 1], scope),
          });
        }
        let total = 0;
        for (let idx = 0; idx < sumRange.length; idx++) {
          const allMatch = pairs.every((p) => matchesCriteria(p.range[idx] ?? null, p.crit));
          if (allMatch) {
            const n = toStrictNumber(sumRange[idx] ?? null);
            if (n !== null) total += n;
          }
        }
        return total;
      }

      case 'COUNTIF': {
        if (args.length !== 2) throw new ExcelError('#VALUE!');
        const range = evalRangeFlat(args[0], scope);
        const criteria = evalToScalar(args[1], scope);
        let count = 0;
        for (const item of range) {
          if (matchesCriteria(item, criteria)) count++;
        }
        return count;
      }

      case 'COUNTIFS': {
        if (args.length < 2 || args.length % 2 !== 0) throw new ExcelError('#VALUE!');
        const pairs: { range: ScalarValue[]; crit: ScalarValue }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          pairs.push({
            range: evalRangeFlat(args[i], scope),
            crit: evalToScalar(args[i + 1], scope),
          });
        }
        const len = pairs[0].range.length;
        let count = 0;
        for (let idx = 0; idx < len; idx++) {
          if (pairs.every((p) => matchesCriteria(p.range[idx] ?? null, p.crit))) {
            count++;
          }
        }
        return count;
      }

      case 'FILTER': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const arr = evalRangeMatrix(args[0], scope);
        const include = evalRangeMatrix(args[1], scope);
        const out: MatrixValue = [];
        for (let r = 0; r < arr.length; r++) {
          const flag = include[r]?.[0] ?? include[0]?.[r] ?? false;
          if (toBoolean(flag)) {
            out.push(arr[r]);
          }
        }
        if (out.length === 0) {
          if (args[2]) return evalNode(args[2], scope) as EvalValue;
          throw new ExcelError('#CALC!');
        }
        return out;
      }

      case 'UNIQUE': {
        const arr = evalRangeMatrix(args[0], scope);
        const seen = new Set<string>();
        const out: MatrixValue = [];
        for (const row of arr) {
          const key = JSON.stringify(row);
          if (!seen.has(key) && row.some((c) => c !== null && c !== '')) {
            seen.add(key);
            out.push(row);
          }
        }
        return out.length ? out : [['']];
      }

      case 'SORT': {
        const arr = evalRangeMatrix(args[0], scope).map((r) => [...r]);
        const sortIdx = args[1] ? Math.trunc(toNumber(evalToScalar(args[1], scope))) - 1 : 0;
        const order = args[2] ? Math.trunc(toNumber(evalToScalar(args[2], scope))) : 1;
        arr.sort((a, b) => {
          const va = a[sortIdx] ?? '';
          const vb = b[sortIdx] ?? '';
          const na = toStrictNumber(va);
          const nb = toStrictNumber(vb);
          if (na !== null && nb !== null) return (na - nb) * order;
          return toStringVal(va).localeCompare(toStringVal(vb)) * order;
        });
        return arr;
      }

      case 'SORTBY': {
        if (args.length < 2) throw new ExcelError('#VALUE!');
        const arr = evalRangeMatrix(args[0], scope).map((r, i) => ({ row: r, idx: i }));
        const byArr = evalRangeFlat(args[1], scope);
        const order = args[2] ? Math.trunc(toNumber(evalToScalar(args[2], scope))) : 1;
        arr.sort((a, b) => {
          const va = byArr[a.idx] ?? '';
          const vb = byArr[b.idx] ?? '';
          const na = toStrictNumber(va);
          const nb = toStrictNumber(vb);
          if (na !== null && nb !== null) return (na - nb) * order;
          return toStringVal(va).localeCompare(toStringVal(vb)) * order;
        });
        return arr.map((x) => x.row);
      }

      case 'SEQUENCE': {
        const rows = Math.max(1, Math.trunc(toNumber(evalToScalar(args[0], scope))));
        const cols = args[1] ? Math.max(1, Math.trunc(toNumber(evalToScalar(args[1], scope)))) : 1;
        const start = args[2] ? toNumber(evalToScalar(args[2], scope)) : 1;
        const step = args[3] ? toNumber(evalToScalar(args[3], scope)) : 1;
        const out: MatrixValue = [];
        let cur = start;
        for (let r = 0; r < rows; r++) {
          const row: ScalarValue[] = [];
          for (let c = 0; c < cols; c++) {
            row.push(cur);
            cur += step;
          }
          out.push(row);
        }
        return out;
      }

      case 'VSTACK': {
        const out: MatrixValue = [];
        for (const a of args) {
          const m = evalRangeMatrix(a, scope);
          for (const row of m) out.push(row);
        }
        return out;
      }

      case 'HSTACK': {
        const matrices = args.map((a) => evalRangeMatrix(a, scope));
        const maxRows = Math.max(...matrices.map((m) => m.length));
        const out: MatrixValue = [];
        for (let r = 0; r < maxRows; r++) {
          const row: ScalarValue[] = [];
          for (const m of matrices) {
            const width = m[0]?.length || 1;
            const rData = m[r] || Array(width).fill('');
            row.push(...rData);
          }
          out.push(row);
        }
        return out;
      }

      case 'TEXTSPLIT': {
        const text = toStringVal(evalToScalar(args[0], scope));
        const colDelim = args[1] ? toStringVal(evalToScalar(args[1], scope)) : ',';
        const parts = text.split(colDelim);
        return [parts];
      }

      case 'GROUPBY':
      case 'PIVOTBY': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const keys = evalRangeFlat(args[0], scope);
        const vals = evalRangeFlat(args[args.length - 2], scope);
        const map = new Map<string, number>();
        for (let i = 0; i < keys.length; i++) {
          const k = toStringVal(keys[i]);
          if (!k) continue;
          const v = toStrictNumber(vals[i] ?? 0) ?? 0;
          map.set(k, (map.get(k) || 0) + v);
        }
        const out: MatrixValue = [];
        for (const [k, sum] of map.entries()) {
          out.push([k, sum]);
        }
        return out.length ? out : [['', 0]];
      }

      case 'LEFT': {
        const text = toStringVal(evalToScalar(args[0], scope));
        const n = args[1] ? Math.max(0, Math.trunc(toNumber(evalToScalar(args[1], scope)))) : 1;
        return text.slice(0, n);
      }

      case 'RIGHT': {
        const text = toStringVal(evalToScalar(args[0], scope));
        const n = args[1] ? Math.max(0, Math.trunc(toNumber(evalToScalar(args[1], scope)))) : 1;
        return n === 0 ? '' : text.slice(-n);
      }

      case 'MID': {
        const text = toStringVal(evalToScalar(args[0], scope));
        const start = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        const count = Math.trunc(toNumber(evalToScalar(args[2], scope)));
        if (start < 1 || count < 0) throw new ExcelError('#VALUE!');
        return text.slice(start - 1, start - 1 + count);
      }

      case 'TRIM': {
        const text = toStringVal(evalToScalar(args[0], scope));
        return text.trim().replace(/\s+/g, ' ');
      }

      case 'CLEAN': {
        const text = toStringVal(evalToScalar(args[0], scope));
        // eslint-disable-next-line no-control-regex
        return text.replace(/[\x00-\x1F]/g, '');
      }

      case 'LEN':
        return toStringVal(evalToScalar(args[0], scope)).length;

      case 'UPPER':
        return toStringVal(evalToScalar(args[0], scope)).toUpperCase();

      case 'LOWER':
        return toStringVal(evalToScalar(args[0], scope)).toLowerCase();

      case 'PROPER':
        return toStringVal(evalToScalar(args[0], scope)).replace(/\w\S*/g, (w) =>
          w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()
        );

      case 'FIND': {
        const needle = toStringVal(evalToScalar(args[0], scope));
        const haystack = toStringVal(evalToScalar(args[1], scope));
        const start = args[2] ? Math.trunc(toNumber(evalToScalar(args[2], scope))) - 1 : 0;
        const idx = haystack.indexOf(needle, start);
        if (idx === -1) throw new ExcelError('#VALUE!');
        return idx + 1;
      }

      case 'SEARCH': {
        const needle = toStringVal(evalToScalar(args[0], scope)).toLowerCase();
        const haystack = toStringVal(evalToScalar(args[1], scope)).toLowerCase();
        const start = args[2] ? Math.trunc(toNumber(evalToScalar(args[2], scope))) - 1 : 0;
        const idx = haystack.indexOf(needle, start);
        if (idx === -1) throw new ExcelError('#VALUE!');
        return idx + 1;
      }

      case 'SUBSTITUTE': {
        const text = toStringVal(evalToScalar(args[0], scope));
        const oldT = toStringVal(evalToScalar(args[1], scope));
        const newT = toStringVal(evalToScalar(args[2], scope));
        return text.split(oldT).join(newT);
      }

      case 'CONCAT':
      case 'CONCATENATE': {
        let res = '';
        for (const a of args) {
          for (const v of evalRangeFlat(a, scope)) {
            res += toStringVal(v);
          }
        }
        return res;
      }

      case 'TEXTJOIN': {
        if (args.length < 3) throw new ExcelError('#VALUE!');
        const delim = toStringVal(evalToScalar(args[0], scope));
        const ignoreEmpty = toBoolean(evalToScalar(args[1], scope));
        const parts: string[] = [];
        for (const a of args.slice(2)) {
          for (const v of evalRangeFlat(a, scope)) {
            const s = toStringVal(v);
            if (!ignoreEmpty || s !== '') parts.push(s);
          }
        }
        return parts.join(delim);
      }

      case 'TEXT': {
        const val = evalToScalar(args[0], scope);
        const fmt = toStringVal(evalToScalar(args[1], scope));
        const num = toStrictNumber(val);
        if (num !== null) {
          if (fmt.includes('%')) {
            const dec = (fmt.split('.')[1] || '').replace(/[^0#]/g, '').length;
            return `${(num * 100).toFixed(dec)}%`;
          }
          if (fmt.includes('$') || fmt.includes('₹')) {
            const sym = fmt.includes('₹') ? '₹' : '$';
            return `${sym}${num.toLocaleString('en-US', { maximumFractionDigits: 2 })}`;
          }
          if (fmt.includes('#,##0')) {
            const dec = (fmt.split('.')[1] || '').replace(/[^0#]/g, '').length;
            return num.toLocaleString('en-US', {
              minimumFractionDigits: dec,
              maximumFractionDigits: dec,
            });
          }
        }
        const dt = parseDateValue(val);
        if (dt) {
          const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          if (fmt.toLowerCase().includes('mmm-yy')) {
            return `${months[dt.getUTCMonth()]}-${String(dt.getUTCFullYear()).slice(-2)}`;
          }
          return formatDateISO(dt);
        }
        return toStringVal(val);
      }

      case 'TODAY': {
        const now = new Date();
        return formatDateISO(new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate())));
      }

      case 'NOW': {
        return new Date().toISOString().slice(0, 16).replace('T', ' ');
      }

      case 'DATE': {
        const y = Math.trunc(toNumber(evalToScalar(args[0], scope)));
        const m = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        const d = Math.trunc(toNumber(evalToScalar(args[2], scope)));
        return formatDateISO(new Date(Date.UTC(y, m - 1, d)));
      }

      case 'YEAR': {
        const dt = parseDateValue(evalToScalar(args[0], scope));
        if (!dt) throw new ExcelError('#VALUE!');
        return dt.getUTCFullYear();
      }

      case 'MONTH': {
        const dt = parseDateValue(evalToScalar(args[0], scope));
        if (!dt) throw new ExcelError('#VALUE!');
        return dt.getUTCMonth() + 1;
      }

      case 'DAY': {
        const dt = parseDateValue(evalToScalar(args[0], scope));
        if (!dt) throw new ExcelError('#VALUE!');
        return dt.getUTCDate();
      }

      case 'EDATE': {
        const dt = parseDateValue(evalToScalar(args[0], scope));
        const months = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        if (!dt) throw new ExcelError('#VALUE!');
        const res = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + months, dt.getUTCDate()));
        return formatDateISO(res);
      }

      case 'EOMONTH': {
        const dt = parseDateValue(evalToScalar(args[0], scope));
        const months = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        if (!dt) throw new ExcelError('#VALUE!');
        const res = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + months + 1, 0));
        return formatDateISO(res);
      }

      case 'DATEDIF': {
        const d1 = parseDateValue(evalToScalar(args[0], scope));
        const d2 = parseDateValue(evalToScalar(args[1], scope));
        const unit = toStringVal(evalToScalar(args[2], scope)).toUpperCase();
        if (!d1 || !d2 || d1 > d2) throw new ExcelError('#NUM!');
        if (unit === 'D') return Math.round((d2.getTime() - d1.getTime()) / 86400000);
        if (unit === 'M') {
          return (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + (d2.getUTCMonth() - d1.getUTCMonth());
        }
        if (unit === 'Y') {
          let years = d2.getUTCFullYear() - d1.getUTCFullYear();
          if (
            d2.getUTCMonth() < d1.getUTCMonth() ||
            (d2.getUTCMonth() === d1.getUTCMonth() && d2.getUTCDate() < d1.getUTCDate())
          ) {
            years--;
          }
          return years;
        }
        throw new ExcelError('#VALUE!');
      }

      case 'NETWORKDAYS': {
        const d1 = parseDateValue(evalToScalar(args[0], scope));
        const d2 = parseDateValue(evalToScalar(args[1], scope));
        if (!d1 || !d2) throw new ExcelError('#VALUE!');
        const holidays = new Set<string>();
        if (args[2]) {
          for (const h of evalRangeFlat(args[2], scope)) {
            const hd = parseDateValue(h);
            if (hd) holidays.add(formatDateISO(hd));
          }
        }
        let count = 0;
        const step = d1 <= d2 ? 1 : -1;
        const cur = new Date(d1.getTime());
        while (step > 0 ? cur <= d2 : cur >= d2) {
          const day = cur.getUTCDay();
          if (day !== 0 && day !== 6 && !holidays.has(formatDateISO(cur))) {
            count += step;
          }
          cur.setUTCDate(cur.getUTCDate() + step);
        }
        return count;
      }

      case 'WORKDAY': {
        const d1 = parseDateValue(evalToScalar(args[0], scope));
        let days = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        if (!d1) throw new ExcelError('#VALUE!');
        const cur = new Date(d1.getTime());
        const step = days >= 0 ? 1 : -1;
        days = Math.abs(days);
        while (days > 0) {
          cur.setUTCDate(cur.getUTCDate() + step);
          const day = cur.getUTCDay();
          if (day !== 0 && day !== 6) days--;
        }
        return formatDateISO(cur);
      }

      case 'PMT': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const nper = toNumber(evalToScalar(args[1], scope));
        const pv = toNumber(evalToScalar(args[2], scope));
        const fv = args[3] ? toNumber(evalToScalar(args[3], scope)) : 0;
        const type = args[4] ? toNumber(evalToScalar(args[4], scope)) : 0;
        return calcPMT(rate, nper, pv, fv, type);
      }

      case 'IPMT': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const per = toNumber(evalToScalar(args[1], scope));
        const nper = toNumber(evalToScalar(args[2], scope));
        const pv = toNumber(evalToScalar(args[3], scope));
        const pmt = calcPMT(rate, nper, pv);
        const fvPrev = calcFV(rate, per - 1, pmt, pv);
        return fvPrev * rate;
      }

      case 'PPMT': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const per = toNumber(evalToScalar(args[1], scope));
        const nper = toNumber(evalToScalar(args[2], scope));
        const pv = toNumber(evalToScalar(args[3], scope));
        const pmt = calcPMT(rate, nper, pv);
        const fvPrev = calcFV(rate, per - 1, pmt, pv);
        return pmt - fvPrev * rate;
      }

      case 'FV': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const nper = toNumber(evalToScalar(args[1], scope));
        const pmt = toNumber(evalToScalar(args[2], scope));
        const pv = args[3] ? toNumber(evalToScalar(args[3], scope)) : 0;
        const type = args[4] ? toNumber(evalToScalar(args[4], scope)) : 0;
        return calcFV(rate, nper, pmt, pv, type);
      }

      case 'PV': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const nper = toNumber(evalToScalar(args[1], scope));
        const pmt = toNumber(evalToScalar(args[2], scope));
        const fv = args[3] ? toNumber(evalToScalar(args[3], scope)) : 0;
        const type = args[4] ? toNumber(evalToScalar(args[4], scope)) : 0;
        return calcPV(rate, nper, pmt, fv, type);
      }

      case 'NPER': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const pmt = toNumber(evalToScalar(args[1], scope));
        const pv = toNumber(evalToScalar(args[2], scope));
        const fv = args[3] ? toNumber(evalToScalar(args[3], scope)) : 0;
        if (rate === 0) return -(pv + fv) / pmt;
        return Math.log((pmt - fv * rate) / (pmt + pv * rate)) / Math.log(1 + rate);
      }

      case 'RATE': {
        const nper = toNumber(evalToScalar(args[0], scope));
        const pmt = toNumber(evalToScalar(args[1], scope));
        const pv = toNumber(evalToScalar(args[2], scope));
        const fv = args[3] ? toNumber(evalToScalar(args[3], scope)) : 0;
        let r = 0.05;
        for (let i = 0; i < 60; i++) {
          const term = Math.pow(1 + r, nper);
          const f = pv * term + pmt * ((term - 1) / r) + fv;
          const f2 = pv * Math.pow(1 + r + 1e-5, nper) + pmt * ((Math.pow(1 + r + 1e-5, nper) - 1) / (r + 1e-5)) + fv;
          const df = (f2 - f) / 1e-5;
          if (Math.abs(f) < 1e-7) return r;
          r -= f / df;
        }
        return r;
      }

      case 'NPV': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const cashflows = getNumericItems(args.slice(1));
        return calcNPV(rate, cashflows);
      }

      case 'XNPV': {
        const rate = toNumber(evalToScalar(args[0], scope));
        const values = evalRangeFlat(args[1], scope).map((v) => toNumber(v));
        const dates = evalRangeFlat(args[2], scope).map((v) => {
          const d = parseDateValue(v);
          if (!d) throw new ExcelError('#VALUE!');
          return d;
        });
        return calcXNPV(rate, values, dates);
      }

      case 'IRR': {
        const values = evalRangeFlat(args[0], scope).map((v) => toNumber(v));
        const guess = args[1] ? toNumber(evalToScalar(args[1], scope)) : 0.1;
        return calcIRR(values, guess);
      }

      case 'XIRR': {
        const values = evalRangeFlat(args[0], scope).map((v) => toNumber(v));
        const dates = evalRangeFlat(args[1], scope).map((v) => {
          const d = parseDateValue(v);
          if (!d) throw new ExcelError('#VALUE!');
          return d;
        });
        const guess = args[2] ? toNumber(evalToScalar(args[2], scope)) : 0.1;
        return calcXIRR(values, dates, guess);
      }

      case 'OFFSET': {
        if (args.length < 3 || args[0].kind !== 'cell') throw new ExcelError('#VALUE!');
        const base = parseCellCoord(args[0].ref);
        if (!base) throw new ExcelError('#REF!');
        const rOff = Math.trunc(toNumber(evalToScalar(args[1], scope)));
        const cOff = Math.trunc(toNumber(evalToScalar(args[2], scope)));
        const height = args[3] ? Math.trunc(toNumber(evalToScalar(args[3], scope))) : 1;
        const width = args[4] ? Math.trunc(toNumber(evalToScalar(args[4], scope))) : 1;
        const startCoord = `${indexToColLetter(base.col + cOff)}${base.row + rOff}`;
        const endCoord = `${indexToColLetter(base.col + cOff + width - 1)}${base.row + rOff + height - 1}`;
        const expanded = expandRange(startCoord, endCoord);
        if (!expanded) throw new ExcelError('#REF!');
        if (height === 1 && width === 1) return getCellEval(startCoord);
        return expanded.coords.map((rowC) => rowC.map((c) => toScalar(getCellEval(c))));
      }

      case 'INDIRECT': {
        const refText = toStringVal(evalToScalar(args[0], scope)).trim();
        if (refText.includes(':')) {
          const [a, b] = refText.split(':');
          const expanded = expandRange(a, b);
          if (!expanded) throw new ExcelError('#REF!');
          return expanded.coords.map((rowC) => rowC.map((c) => toScalar(getCellEval(c))));
        }
        const parsed = parseCellCoord(refText);
        if (!parsed) throw new ExcelError('#REF!');
        return getCellEval(parsed.coord);
      }

      default:
        throw new ExcelError('#NAME?');
    }
  }

  // First pass: evaluate all defined cells
  const spillQueue: { origin: string; matrix: MatrixValue; format?: CellFormat; currencySymbol?: string }[] = [];

  for (const [coord, cell] of Object.entries(cells)) {
    const isFormula = cell.raw.trim().startsWith('=');
    try {
      const val = getCellEval(coord);
      if (isMatrix(val)) {
        const topVal = val[0]?.[0] ?? null;
        computed[coord] = {
          raw: cell.raw,
          value: topVal,
          display: formatCellValue(topVal, cell.format, cell.currencySymbol),
          isFormula,
          format: cell.format,
          currencySymbol: cell.currencySymbol,
        };
        spillQueue.push({
          origin: coord,
          matrix: val,
          format: cell.format,
          currencySymbol: cell.currencySymbol,
        });
      } else if (isLambda(val)) {
        computed[coord] = {
          raw: cell.raw,
          value: '#CALC!',
          display: '#CALC!',
          isFormula,
          error: '#CALC!',
        };
      } else {
        computed[coord] = {
          raw: cell.raw,
          value: val,
          display: formatCellValue(val, cell.format, cell.currencySymbol),
          isFormula,
          format: cell.format,
          currencySymbol: cell.currencySymbol,
        };
      }
    } catch (err) {
      const code = err instanceof ExcelError ? err.code : '#ERROR!';
      computed[coord] = {
        raw: cell.raw,
        value: code,
        display: code,
        isFormula,
        error: code,
        format: cell.format,
        currencySymbol: cell.currencySymbol,
      };
    }
  }

  // Second pass: spill matrix outputs into empty cells
  for (const spill of spillQueue) {
    const originParsed = parseCellCoord(spill.origin);
    if (!originParsed) continue;
    for (let r = 0; r < spill.matrix.length; r++) {
      for (let c = 0; c < spill.matrix[r].length; c++) {
        if (r === 0 && c === 0) continue;
        const targetCoord = `${indexToColLetter(originParsed.col + c)}${originParsed.row + r}`;
        const existing = cells[targetCoord];
        if (existing && existing.raw.trim() !== '') continue;
        const val = spill.matrix[r][c];
        computed[targetCoord] = {
          raw: '',
          value: val,
          display: formatCellValue(val, spill.format, spill.currencySymbol),
          isFormula: false,
          isSpilled: true,
          spillSource: spill.origin,
          format: spill.format,
          currencySymbol: spill.currencySymbol,
        };
      }
    }
  }

  return computed;
}

export function gradeFormulaAnswer(
  userFormula: string,
  question: TestQuestion,
  baseCells: Record<string, CellData>
): GradeResult {
  const trimmed = userFormula.trim();
  const warnings: string[] = [];
  const details: string[] = [];

  if (!trimmed) {
    return {
      passed: false,
      scoreAwarded: 0,
      maxPoints: question.points,
      computedDisplay: '(empty)',
      feedbackTitle: 'No formula entered',
      feedbackDetails: [`Click cell ${question.targetCell} or use the formula bar to write an Excel formula starting with '='.`],
      bestPracticeWarnings: [],
    };
  }

  if (!trimmed.startsWith('=')) {
    return {
      passed: false,
      scoreAwarded: 0,
      maxPoints: question.points,
      computedDisplay: trimmed,
      feedbackTitle: 'Missing "=" formula prefix',
      feedbackDetails: [
        'You entered a static value instead of a dynamic Excel formula.',
        'Begin your formula with "=" and reference the input cells so updates recalculate automatically.',
      ],
      bestPracticeWarnings: ['Static values fail audit checks when source inputs change.'],
    };
  }

  const testCells: Record<string, CellData> = {
    ...baseCells,
    [question.targetCell]: {
      ...(baseCells[question.targetCell] || {}),
      raw: trimmed,
    },
  };

  const evaluated = evaluateSheet(testCells);
  const targetComputed = evaluated[question.targetCell];

  if (!targetComputed || targetComputed.error) {
    return {
      passed: false,
      scoreAwarded: 0,
      maxPoints: question.points,
      computedDisplay: targetComputed?.display || '#ERROR!',
      feedbackTitle: `Formula returned ${targetComputed?.error || '#ERROR!'}`,
      feedbackDetails: [
        `Check argument syntax, cell coordinates, and parentheses in ${question.targetCell}.`,
      ],
      bestPracticeWarnings: [],
    };
  }

  let valueMatches = false;
  if (typeof question.expectedValue === 'number') {
    const tol = question.numericTolerance ?? 0.01;
    const actualNum = toStrictNumber(targetComputed.value);
    if (actualNum !== null && Math.abs(actualNum - question.expectedValue) <= tol) {
      valueMatches = true;
    }
  } else {
    const expectedStr = String(question.expectedValue).trim().toLowerCase();
    const actualStr = toStringVal(targetComputed.value).trim().toLowerCase();
    if (expectedStr === actualStr) {
      valueMatches = true;
    }
  }

  if (!valueMatches) {
    details.push(
      `Your formula produced "${targetComputed.display}", whereas the expected result is "${formatCellValue(question.expectedValue, baseCells[question.targetCell]?.format, baseCells[question.targetCell]?.currencySymbol)}".`
    );
  } else {
    details.push(`Output "${targetComputed.display}" matches the expected business calculation.`);
  }

  if (question.forbiddenHardcodedNumbers) {
    for (const num of question.forbiddenHardcodedNumbers) {
      const numRegex = new RegExp(`(?<![A-Za-z0-9.])${num}(?![0-9.])`);
      if (numRegex.test(trimmed)) {
        warnings.push(`Avoid hardcoding ${num} directly inside the formula; reference the assumption cell instead.`);
      }
    }
  }

  if (question.requiredPatterns) {
    for (const rule of question.requiredPatterns) {
      const re = new RegExp(rule.regex, 'i');
      if (!re.test(trimmed)) {
        warnings.push(rule.message);
      }
    }
  }

  const passed = valueMatches && warnings.length === 0;
  const partialScore = valueMatches && warnings.length > 0 ? Math.round(question.points * 0.6) : 0;

  return {
    passed,
    scoreAwarded: passed ? question.points : partialScore,
    maxPoints: question.points,
    computedDisplay: targetComputed.display,
    feedbackTitle: passed
      ? 'Clean Formula & Best Practices'
      : valueMatches
      ? 'Result Matches, Needs Best-Practice Fix'
      : 'Incorrect Output — Review Formula Logic',
    feedbackDetails: details,
    bestPracticeWarnings: warnings,
  };
}
