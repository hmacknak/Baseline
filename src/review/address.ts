// A1 address helpers. Rows and columns are 0-based internally.

export function colToLetters(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function lettersToCol(letters: string): number {
  let n = 0;
  const up = letters.toUpperCase();
  for (let i = 0; i < up.length; i++) {
    n = n * 26 + (up.charCodeAt(i) - 64);
  }
  return n - 1;
}

export function toA1(row: number, col: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

/** Parse "B7" or "$B$7" into 0-based coordinates. Returns null if not a single-cell address. */
export function parseA1(address: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(address.trim());
  if (!m) return null;
  return { row: parseInt(m[2], 10) - 1, col: lettersToCol(m[1]) };
}
