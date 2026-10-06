import { parseA1, toA1 } from "../src/review/address";
import { Cell, SheetSnapshot } from "../src/review/types";

/**
 * Build a snapshot from { A1: value } pairs.
 * Strings starting with "=" are formulas (value treated as a number), strings starting with "#" are errors.
 */
export function sheet(cells: Record<string, string | number | boolean>, name = "Sheet1"): SheetSnapshot {
  let maxRow = 0;
  let maxCol = 0;
  const parsed = Object.keys(cells).map(addr => {
    const p = parseA1(addr)!;
    maxRow = Math.max(maxRow, p.row);
    maxCol = Math.max(maxCol, p.col);
    return { ...p, raw: cells[addr] };
  });
  const grid: Cell[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    grid.push([]);
    for (let c = 0; c <= maxCol; c++) {
      grid[r].push({ row: r, col: c, address: toA1(r, c), kind: "empty", value: "", formula: null });
    }
  }
  for (const p of parsed) {
    const cell = grid[p.row][p.col];
    if (typeof p.raw === "string" && p.raw.charAt(0) === "=") {
      cell.formula = p.raw;
      cell.kind = "number";
      cell.value = 0;
    } else if (typeof p.raw === "string" && p.raw.charAt(0) === "#") {
      cell.kind = "error";
      cell.value = p.raw;
    } else if (typeof p.raw === "number") {
      cell.kind = "number";
      cell.value = p.raw;
    } else if (typeof p.raw === "boolean") {
      cell.kind = "boolean";
      cell.value = p.raw;
    } else {
      cell.kind = "text";
      cell.value = p.raw;
    }
  }
  return { name, cells: grid, originRow: 0, originCol: 0 };
}
