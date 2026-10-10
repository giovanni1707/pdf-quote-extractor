/**
 * Utilities for adjusting Excel formula references when inserting rows
 */

export function colToNum(col: string): number {
  let num = 0;
  const clean = col.toUpperCase().replace(/\$/g, '');
  for (let i = 0; i < clean.length; i++) {
    num = num * 26 + (clean.charCodeAt(i) - 64);
  }
  return num;
}

export function numToCol(num: number): string {
  let s = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    num = Math.floor((num - mod) / 26);
  }
  return s;
}

/**
 * Adjusts cell references within an Excel formula when new item rows are inserted
 */
export function shiftFormula(
  formula: string,
  options: {
    isNewItemRow?: boolean;
    sourceRow?: number;
    targetRow?: number;
    firstItemRow: number;
    oldLastItemRow: number;
    delta: number;
    oldTotalsRow: number;
  }
): string {
  if (!formula || typeof formula !== 'string') return '';
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

  const { isNewItemRow, sourceRow, targetRow, firstItemRow, oldLastItemRow, delta, oldTotalsRow } = options;
  const newLastItemRow = oldLastItemRow + delta;
  const newTotalsRow = oldTotalsRow + delta;

  // 1. If it's a new item row copied from sourceRow:
  if (isNewItemRow && sourceRow && targetRow) {
    // Replace references to sourceRow with targetRow
    // Matches patterns like $?A$?7, $?BC$?7
    const refRegex = /([$]?[A-Za-z]{1,3})([$]?)(\d+)/g;
    return cleanFormula.replace(refRegex, (match, colPart, dollarRow, rowStr) => {
      const r = parseInt(rowStr, 10);
      // If reference points to the source item row, point to targetRow
      if (r === sourceRow) {
        return `${colPart}${dollarRow}${targetRow}`;
      }
      // If reference points to a row >= oldTotalsRow, shift by delta
      if (r >= oldTotalsRow) {
        return `${colPart}${dollarRow}${r + delta}`;
      }
      return match;
    });
  }

  // 2. If it's the Totals row or a row that moved down:
  // First, check for range patterns like G7:G16 or $G$7:$G$16
  const rangeRegex = /([$]?[A-Za-z]{1,3})([$]?)(\d+):([$]?[A-Za-z]{1,3})([$]?)(\d+)/g;
  let updated = cleanFormula.replace(rangeRegex, (match, col1, d1, r1Str, col2, d2, r2Str) => {
    let r1 = parseInt(r1Str, 10);
    let r2 = parseInt(r2Str, 10);

    // If range started at firstItemRow and ended at oldLastItemRow:
    if (r1 === firstItemRow && r2 === oldLastItemRow) {
      return `${col1}${d1}${r1}:${col2}${d2}${newLastItemRow}`;
    }

    // If both r1 and r2 are >= oldTotalsRow, shift both
    if (r1 >= oldTotalsRow && r2 >= oldTotalsRow) {
      return `${col1}${d1}${r1 + delta}:${col2}${d2}${r2 + delta}`;
    }

    return match;
  });

  // Next, single cell references
  const singleRefRegex = /(?<![:A-Za-z0-9_])([$]?[A-Za-z]{1,3})([$]?)(\d+)(?![:A-Za-z0-9_])/g;
  updated = updated.replace(singleRefRegex, (match, colPart, dollarRow, rowStr) => {
    const r = parseInt(rowStr, 10);
    if (r >= oldTotalsRow) {
      return `${colPart}${dollarRow}${r + delta}`;
    }
    return match;
  });

  return updated;
}
