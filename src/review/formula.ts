// A small, forgiving lexer for Excel A1 formulas. It understands enough of the
// grammar to find cell references, numeric literals and external links; it is
// not a full parser. Written without regex lookbehind so it runs in older
// Office webviews.

import { colToLetters, lettersToCol } from "./address";

export type TokenType =
  | "string" // "text"
  | "sheet" // Sheet1!  or  'My Sheet'!
  | "bracket" // [anything] – structured refs or external workbook index
  | "func" // SUM  (immediately followed by "(")
  | "ref" // A1, $B$2
  | "col" // A   (only meaningful next to ":")
  | "number" // 12, 1.05  (also row numbers in 1:1 ranges)
  | "name" // defined names, TRUE, FALSE
  | "op" // + - * / ^ & = < > , ( ) : ; % { }
  | "space";

export interface Token {
  type: TokenType;
  text: string;
}

const WORD_CHAR = /[A-Za-z0-9_.$\\?]/;
const REF = /^\$?[A-Za-z]{1,3}\$?\d+$/;
const COL = /^\$?[A-Za-z]{1,3}$/;
const NUMBER = /^\d+(\.\d+)?$|^\.\d+$/;

export function lex(formula: string): Token[] {
  const src = formula.charAt(0) === "=" ? formula.slice(1) : formula;
  const tokens: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src.charAt(i);

    if (ch === '"') {
      let j = i + 1;
      while (j < src.length) {
        if (src.charAt(j) === '"') {
          if (src.charAt(j + 1) === '"') {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      tokens.push({ type: "string", text: src.slice(i, j + 1) });
      i = j + 1;
      continue;
    }

    if (ch === "'") {
      let j = i + 1;
      while (j < src.length) {
        if (src.charAt(j) === "'") {
          if (src.charAt(j + 1) === "'") {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      let end = j + 1;
      if (src.charAt(end) === "!") end++;
      tokens.push({ type: "sheet", text: src.slice(i, end) });
      i = end;
      continue;
    }

    if (ch === "[") {
      let depth = 0;
      let j = i;
      while (j < src.length) {
        const c = src.charAt(j);
        if (c === "[") depth++;
        else if (c === "]") {
          depth--;
          if (depth === 0) break;
        }
        j++;
      }
      tokens.push({ type: "bracket", text: src.slice(i, j + 1) });
      i = j + 1;
      continue;
    }

    if (ch === " " || ch === "\n" || ch === "\r" || ch === "\t") {
      let j = i;
      while (j < src.length && /\s/.test(src.charAt(j))) j++;
      tokens.push({ type: "space", text: src.slice(i, j) });
      i = j;
      continue;
    }

    if (WORD_CHAR.test(ch)) {
      let j = i;
      while (j < src.length && WORD_CHAR.test(src.charAt(j))) j++;
      const word = src.slice(i, j);
      const next = src.charAt(j);
      if (next === "!") {
        tokens.push({ type: "sheet", text: word + "!" });
        i = j + 1;
        continue;
      }
      if (next === "(") tokens.push({ type: "func", text: word });
      else if (NUMBER.test(word)) tokens.push({ type: "number", text: word });
      else if (REF.test(word)) tokens.push({ type: "ref", text: word });
      else if (COL.test(word)) tokens.push({ type: "col", text: word });
      else tokens.push({ type: "name", text: word });
      i = j;
      continue;
    }

    tokens.push({ type: "op", text: ch });
    i++;
  }
  return tokens;
}

function prevSolid(tokens: Token[], i: number): Token | undefined {
  for (let k = i - 1; k >= 0; k--) if (tokens[k].type !== "space") return tokens[k];
  return undefined;
}

function nextSolid(tokens: Token[], i: number): Token | undefined {
  for (let k = i + 1; k < tokens.length; k++) if (tokens[k].type !== "space") return tokens[k];
  return undefined;
}

function prevSolidIndex(tokens: Token[], i: number): number {
  for (let k = i - 1; k >= 0; k--) if (tokens[k].type !== "space") return k;
  return -1;
}

const ARITH = ["+", "-", "*", "/", "^"];

/**
 * Numeric literals used as arithmetic operands, e.g. 1.05 in "=A1*1.05" or 500 in "=SUM(A1:A3)+500".
 * Function arguments such as the 2 in ROUND(A1,2) or the 3 in VLOOKUP(...,3,FALSE) are ignored,
 * as are 0 and 1 and anything inside array constants.
 */
export function findHardcodedNumbers(formula: string): string[] {
  const tokens = lex(formula);
  const found: string[] = [];
  let braceDepth = 0;
  const solid = tokens.filter(t => t.type !== "space");

  // A formula that is nothing but a number ("=500") is a hard-code by definition.
  if (solid.length === 1 && solid[0].type === "number") {
    return Number(solid[0].text) === 0 || Number(solid[0].text) === 1 ? [] : [solid[0].text];
  }

  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === "op" && t.text === "{") braceDepth++;
    if (t.type === "op" && t.text === "}") braceDepth--;
    if (t.type !== "number" || braceDepth > 0) continue;

    const value = Number(t.text);
    if (value === 0 || value === 1) continue;

    const prev = prevSolid(tokens, i);
    const next = nextSolid(tokens, i);
    // Row ranges like 1:1 or 5:10.
    if ((prev && prev.text === ":") || (next && next.text === ":")) continue;

    let prevIsOperator = !!prev && prev.type === "op" && ARITH.indexOf(prev.text) >= 0;
    if (prevIsOperator && prev && prev.text === "-") {
      // Unary minus directly after "(" "," or "=" is a sign, not subtraction: ROUND(A1,-2).
      const before = tokens[prevSolidIndex(tokens, prevSolidIndex(tokens, i))];
      if (!before || (before.type === "op" && (before.text === "(" || before.text === ","))) {
        prevIsOperator = false;
      }
    }
    const nextIsOperator =
      !!next && next.type === "op" && (ARITH.indexOf(next.text) >= 0 || next.text === "%");
    const percentThenOperator =
      !!next &&
      next.text === "%" &&
      (() => {
        const after = nextSolid(tokens, tokens.indexOf(next));
        return !!after && after.type === "op" && ARITH.indexOf(after.text) >= 0;
      })();

    if (prevIsOperator || (nextIsOperator && next!.text !== "%") || percentThenOperator) {
      found.push(next && next.text === "%" ? t.text + "%" : t.text);
    }
  }
  return found;
}

/**
 * True when the formula points at another workbook, e.g. "='[Budget.xlsx]Sheet1'!A1" or "=[1]Sheet1!A1".
 */
export function hasExternalLink(formula: string): boolean {
  const tokens = lex(formula).filter(t => t.type !== "space");
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === "sheet" && t.text.charAt(0) === "'" && t.text.indexOf("[") >= 0) return true;
    if (t.type === "bracket") {
      const prev = tokens[i - 1];
      const next = tokens[i + 1];
      const startsExpression = !prev || prev.type === "op";
      if (startsExpression && next && (next.type === "sheet" || next.type === "name")) return true;
      if (/\.(xl[a-z]*|csv)\]$/i.test(t.text)) return true;
    }
  }
  return false;
}

function refToR1C1(ref: string, row: number, col: number): string {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref)!;
  const c = lettersToCol(m[2]);
  const r = parseInt(m[4], 10) - 1;
  const rPart = m[3] ? `R${r + 1}` : r === row ? "R" : `R[${r - row}]`;
  const cPart = m[1] ? `C${c + 1}` : c === col ? "C" : `C[${c - col}]`;
  return rPart + cPart;
}

function colToR1C1(text: string, col: number): string {
  const abs = text.charAt(0) === "$";
  const c = lettersToCol(abs ? text.slice(1) : text);
  return abs ? `C${c + 1}` : c === col ? "C" : `C[${c - col}]`;
}

function rowToR1C1(text: string, row: number): string {
  const abs = text.charAt(0) === "$";
  const r = parseInt(abs ? text.slice(1) : text, 10) - 1;
  return abs ? `R${r + 1}` : r === row ? "R" : `R[${r - row}]`;
}

/**
 * Convert an A1 formula to a normalised relative form so formulas copied across a row or column
 * compare equal. Mirrors Excel's own R1C1 notation closely enough for consistency checks.
 */
export function toR1C1(formula: string, row: number, col: number): string {
  const tokens = lex(formula);
  const out: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    const prev = prevSolid(tokens, i);
    const next = nextSolid(tokens, i);
    const nearColon = (prev && prev.text === ":") || (next && next.text === ":");
    if (t.type === "space") continue;
    if (t.type === "ref") out.push(refToR1C1(t.text, row, col));
    else if (t.type === "col" && nearColon) out.push(colToR1C1(t.text, col));
    else if (t.type === "number" && nearColon && /^\d+$/.test(t.text)) out.push(rowToR1C1(t.text, row));
    else if (t.type === "string") out.push(t.text);
    else out.push(t.text.toUpperCase());
  }
  return "=" + out.join("");
}

export { colToLetters };
