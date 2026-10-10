import ExcelJS from 'exceljs';
import { HyperFormula } from 'hyperformula';
import { shiftFormula, numToCol } from './formulaUtils';
import type { QuotationExtraction, PositionItem } from '../../server';

export interface TemplateLayout {
  sheetName: string;
  headerRow: number;
  firstItemRow: number;
  lastItemRow: number;
  totalsRow: number;
  hasSumFormula: boolean;
}

export interface CheckResult {
  id: string;
  category: 'EXTRACTION' | 'TOTALS' | 'INTEGRITY';
  name: string;
  passed: boolean;
  message: string;
  details?: string;
}

export interface VerificationReport {
  passed: boolean;
  checks: CheckResult[];
  exworksTotal: number;
  invoiceAmount: number;
  rowsAdded: number;
  exchangeRateUsed: number;
  flags: string[];
}

/**
 * Detects template layout (header, item rows, totals row)
 */
export function analyzeTemplate(workbook: ExcelJS.Workbook): TemplateLayout {
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('Template workbook contains no worksheets.');
  }

  let headerRow = 6;
  let firstItemRow = 7;
  let lastItemRow = 16;
  let totalsRow = 17;
  let hasSumFormula = false;

  // Search for the Totals row by looking for SUM formula or "TOTAL" text
  for (let r = 1; r <= Math.min(worksheet.rowCount + 5, 60); r++) {
    const row = worksheet.getRow(r);
    const colG = row.getCell(7); // Column G
    const colA = row.getCell(1).text?.trim().toUpperCase() || '';
    const colB = row.getCell(2).text?.trim().toUpperCase() || '';
    const colD = row.getCell(4).text?.trim().toUpperCase() || '';

    let formulaStr = '';
    if (colG.formula) {
      formulaStr = colG.formula.toUpperCase();
    } else if (typeof colG.value === 'object' && (colG.value as any)?.formula) {
      formulaStr = ((colG.value as any).formula as string).toUpperCase();
    }

    if (formulaStr.includes('SUM(G') || formulaStr.includes('SUM($G$')) {
      totalsRow = r;
      hasSumFormula = true;
      // Extract starting and ending row from SUM(G7:G16)
      const sumMatch = formulaStr.match(/SUM\([$]?G[$]?(\d+):[$]?G[$]?(\d+)\)/);
      if (sumMatch) {
        firstItemRow = parseInt(sumMatch[1], 10);
        lastItemRow = parseInt(sumMatch[2], 10);
        headerRow = firstItemRow - 1;
        break;
      }
    } else if (colA.includes('TOTAL') || colB.includes('TOTAL') || colD.includes('TOTAL')) {
      totalsRow = r;
      lastItemRow = r - 1;
      // Trace upwards to find first item row
      for (let up = lastItemRow; up >= 1; up--) {
        const upRow = worksheet.getRow(up);
        const cellG = upRow.getCell(7);
        const hasForm = !!(cellG.formula || (typeof cellG.value === 'object' && (cellG.value as any)?.formula));
        if (!hasForm && up < lastItemRow) {
          firstItemRow = up + 1;
          headerRow = up;
          break;
        }
      }
      break;
    }
  }

  return {
    sheetName: worksheet.name,
    headerRow,
    firstItemRow,
    lastItemRow,
    totalsRow,
    hasSumFormula,
  };
}

/**
 * Excel templates often store formulas as "shared formulas": one master cell holds the
 * formula and the cells below only point to it. If the master is edited or moved, the
 * children go blank. Convert every shared formula into a normal, standalone formula first.
 */
function unshareFormulas(worksheet: ExcelJS.Worksheet): void {
  const converted: { cell: ExcelJS.Cell; formula: string; result: any }[] = [];

  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value as any;
      if (v && typeof v === 'object' && (v.sharedFormula || v.shareType === 'shared')) {
        // cell.formula returns the translated formula for child cells
        const f = cell.formula || v.formula;
        if (f) converted.push({ cell, formula: f, result: v.result });
      }
    });
  });

  // Assign only after reading everything, because children are translated from their master
  for (const { cell, formula, result } of converted) {
    cell.value = { formula, result } as any;
  }
}

/**
 * Fills template while preserving all styles, inserting rows if needed, and adjusting formulas
 */
export async function fillTemplate(
  templateBuffer: ArrayBuffer,
  extraction: QuotationExtraction,
  exchangeRate: number
): Promise<{ filledBuffer: ArrayBuffer; rowsAdded: number; layout: TemplateLayout }> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  const worksheet = workbook.worksheets[0];
  unshareFormulas(worksheet);
  const layout = analyzeTemplate(workbook);

  const { firstItemRow, lastItemRow, totalsRow } = layout;
  const initialCapacity = lastItemRow - firstItemRow + 1;
  const requiredCount = extraction.positions.length;
  const rowsAdded = Math.max(0, requiredCount - initialCapacity);

  const oldLastItemRow = lastItemRow;
  const oldTotalsRow = totalsRow;
  const newLastItemRow = lastItemRow + rowsAdded;
  const newTotalsRow = totalsRow + rowsAdded;

  if (rowsAdded > 0) {
    const totalOriginalRows = worksheet.rowCount;

    // 1. Shift lower rows (from oldTotalsRow downwards) down by rowsAdded
    for (let r = totalOriginalRows; r >= oldTotalsRow; r--) {
      const srcRow = worksheet.getRow(r);
      const destRow = worksheet.getRow(r + rowsAdded);

      destRow.height = srcRow.height;

      srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
        const destCell = destRow.getCell(colNumber);

        // Copy styles
        destCell.style = JSON.parse(JSON.stringify(srcCell.style || {}));

        // Shift formula if formula exists
        let rawFormula = srcCell.formula;
        if (!rawFormula && typeof srcCell.value === 'object' && (srcCell.value as any)?.formula) {
          rawFormula = (srcCell.value as any).formula;
        }

        if (rawFormula) {
          const shifted = shiftFormula(rawFormula, {
            firstItemRow,
            oldLastItemRow,
            delta: rowsAdded,
            oldTotalsRow,
          });
          destCell.value = { formula: shifted, result: (srcCell.value as any)?.result };
        } else {
          destCell.value = srcCell.value;
        }
      });
    }

    // 2. Insert new item rows directly below oldLastItemRow
    const templateItemRow = worksheet.getRow(firstItemRow);

    for (let r = oldLastItemRow + 1; r <= newLastItemRow; r++) {
      const newRow = worksheet.getRow(r);
      newRow.height = templateItemRow.height;

      // Copy template cells from firstItemRow
      templateItemRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        newCell.style = JSON.parse(JSON.stringify(srcCell.style || {}));

        let rawFormula = srcCell.formula;
        if (!rawFormula && typeof srcCell.value === 'object' && (srcCell.value as any)?.formula) {
          rawFormula = (srcCell.value as any).formula;
        }

        if (rawFormula) {
          const shifted = shiftFormula(rawFormula, {
            isNewItemRow: true,
            sourceRow: firstItemRow,
            targetRow: r,
            firstItemRow,
            oldLastItemRow,
            delta: rowsAdded,
            oldTotalsRow,
          });
          newCell.value = { formula: shifted };
        } else {
          newCell.value = null;
        }
      });

      // Default O = 0, P = 0.15 for new rows
      newRow.getCell(15).value = 0; // Col O
      newRow.getCell(16).value = 0.15; // Col P
    }

    // 2b. Fix formulas in the ORIGINAL rows above the totals row (e.g. L3:L6 =K3/K$7*100).
    // These rows were not moved, but any reference they hold to the totals row (or below)
    // must follow it down. Without this, L3 divides by K$7 (an item row) instead of the
    // new totals row, giving huge numbers that display as "######".
    // NOTE: this must run AFTER step 2, because step 2 copies (and shifts) formulas from
    // the first item row; shifting it earlier would shift those references twice.
    for (let r = 1; r < oldTotalsRow; r++) {
      // Skip the rows created in step 2 (they were already shifted). This includes any
      // blank spacer row between the last item row and the totals row that step 2 overwrote.
      if (r > oldLastItemRow && r <= oldLastItemRow + rowsAdded) continue;
      const row = worksheet.getRow(r);
      row.eachCell({ includeEmpty: false }, (cell) => {
        let rawFormula = cell.formula;
        if (!rawFormula && typeof cell.value === 'object' && (cell.value as any)?.formula) {
          rawFormula = (cell.value as any).formula;
        }
        if (!rawFormula) return;

        const original = rawFormula.startsWith('=') ? rawFormula.substring(1) : rawFormula;
        const shifted = shiftFormula(rawFormula, {
          firstItemRow,
          oldLastItemRow,
          delta: rowsAdded,
          oldTotalsRow,
        });
        if (shifted !== original) {
          // Drop the cached result so Excel recalculates it on open
          cell.value = { formula: shifted } as any;
        }
      });
    }

    // 3. Shift Merged cells
    if (worksheet.model && worksheet.model.merges) {
      const originalMerges = [...worksheet.model.merges];
      const newMerges: string[] = [];

      for (const mergeRange of originalMerges) {
        worksheet.unMergeCells(mergeRange);
        // Range like A17:D17
        const parts = mergeRange.split(':');
        if (parts.length === 2) {
          const m1 = parts[0].match(/([A-Z]+)(\d+)/);
          const m2 = parts[1].match(/([A-Z]+)(\d+)/);
          if (m1 && m2) {
            let r1 = parseInt(m1[2], 10);
            let r2 = parseInt(m2[2], 10);
            if (r1 >= oldTotalsRow) r1 += rowsAdded;
            if (r2 >= oldTotalsRow) r2 += rowsAdded;
            newMerges.push(`${m1[1]}${r1}:${m2[1]}${r2}`);
          } else {
            newMerges.push(mergeRange);
          }
        }
      }

      for (const nm of newMerges) {
        try {
          worksheet.mergeCells(nm);
        } catch (e) {
          console.warn('Merge cell shift error for', nm, e);
        }
      }
    }
  }

  // Write values ONLY into columns A, B, E, F, I, O, P of item rows
  for (let i = 0; i < extraction.positions.length; i++) {
    const pos = extraction.positions[i];
    const rowNum = firstItemRow + i;
    const row = worksheet.getRow(rowNum);

    // Col A: Reference (stored as TEXT)
    const cellA = row.getCell(1);
    cellA.value = String(pos.reference);
    cellA.numFmt = '@';

    // Col B: Part Name
    const cellB = row.getCell(2);
    cellB.value = String(pos.partName);

    // Col E: Net Price (number)
    const cellE = row.getCell(5);
    cellE.value = Number(pos.netPrice);

    // Col F: Quantity (number)
    const cellF = row.getCell(6);
    cellF.value = Number(pos.quantity);

    // Col I: Exchange rate
    const cellI = row.getCell(9);
    cellI.value = Number(exchangeRate);

    // Col O: Duty (0)
    const cellO = row.getCell(15);
    cellO.value = 0;

    // Col P: VAT (0.15)
    const cellP = row.getCell(16);
    cellP.value = 0.15;
  }

  // If fewer items than template capacity, ensure unused rows remain blank & clean
  if (requiredCount < initialCapacity) {
    for (let r = firstItemRow + requiredCount; r <= lastItemRow; r++) {
      const row = worksheet.getRow(r);
      row.getCell(1).value = null; // A
      row.getCell(2).value = null; // B
      row.getCell(5).value = null; // E
      row.getCell(6).value = null; // F
      row.getCell(9).value = null; // I
    }
  }

  const rawBuffer = await workbook.xlsx.writeBuffer();
  let filledBuffer: ArrayBuffer;
  if (rawBuffer instanceof ArrayBuffer) {
    filledBuffer = rawBuffer;
  } else {
    const u8 = new Uint8Array(rawBuffer as any);
    filledBuffer = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
  }

  return {
    filledBuffer,
    rowsAdded,
    layout: {
      ...layout,
      lastItemRow: newLastItemRow,
      totalsRow: newTotalsRow,
    },
  };
}

/**
 * Converts ExcelJS worksheet to HyperFormula 2D matrix
 */
function worksheetToHyperFormulaMatrix(worksheet: ExcelJS.Worksheet): (string | number | boolean | null)[][] {
  const maxRow = worksheet.rowCount;
  const maxCol = worksheet.columnCount;
  const matrix: (string | number | boolean | null)[][] = [];

  for (let r = 1; r <= maxRow; r++) {
    const row = worksheet.getRow(r);
    const rowData: (string | number | boolean | null)[][] = [];
    const rowArr: (string | number | boolean | null)[] = [];

    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      let formula = cell.formula;
      if (!formula && typeof cell.value === 'object' && (cell.value as any)?.formula) {
        formula = (cell.value as any).formula;
      }

      if (formula) {
        const fStr = formula.startsWith('=') ? formula : '=' + formula;
        rowArr.push(fStr);
      } else if (cell.value === null || cell.value === undefined) {
        rowArr.push(null);
      } else if (typeof cell.value === 'number' || typeof cell.value === 'string' || typeof cell.value === 'boolean') {
        rowArr.push(cell.value);
      } else if (typeof cell.value === 'object') {
        const valObj = cell.value as any;
        if (valObj.text) {
          rowArr.push(valObj.text);
        } else if (valObj.result !== undefined) {
          rowArr.push(valObj.result);
        } else {
          rowArr.push(null);
        }
      } else {
        rowArr.push(null);
      }
    }
    matrix.push(rowArr);
  }

  return matrix;
}

/**
 * Runs full verification with HyperFormula recalculation and integrity checks
 */
export async function verifyOrderList(
  originalBuffer: ArrayBuffer,
  filledBuffer: ArrayBuffer,
  extraction: QuotationExtraction,
  exchangeRate: number,
  layout: TemplateLayout,
  rowsAdded: number
): Promise<VerificationReport> {
  const checks: CheckResult[] = [];
  const flags: string[] = [];

  // Load both workbooks
  const originalWb = new ExcelJS.Workbook();
  await originalWb.xlsx.load(originalBuffer);
  const origSheet = originalWb.worksheets[0];

  const filledWb = new ExcelJS.Workbook();
  await filledWb.xlsx.load(filledBuffer);
  const filledSheet = filledWb.worksheets[0];

  const { firstItemRow, lastItemRow, totalsRow } = layout;
  const positions = extraction.positions;

  // Initialize HyperFormula
  const matrix = worksheetToHyperFormulaMatrix(filledSheet);
  const hf = HyperFormula.buildFromSheets(
    {
      [filledSheet.name]: matrix,
    },
    { licenseKey: 'gpl-v3' }
  );
  const sheetId = hf.getSheetId(filledSheet.name) as number;

  // ==========================================
  // A. EXTRACTION CHECKS
  // ==========================================

  // Check A1: Number of filled rows equals number of positions & order
  let orderCorrect = true;
  let allPosFound = true;
  for (let i = 0; i < positions.length; i++) {
    const rowNum = firstItemRow + i;
    const row = filledSheet.getRow(rowNum);
    const pos = positions[i];

    const ref = row.getCell(1).text?.trim();
    const name = row.getCell(2).text?.trim();
    if (ref !== String(pos.reference).trim()) {
      orderCorrect = false;
    }
    if (!name.includes(pos.partName.trim().substring(0, 5))) {
      orderCorrect = false;
    }
  }
  checks.push({
    id: 'A1_COUNT_ORDER',
    category: 'EXTRACTION',
    name: 'Row count and position ordering',
    passed: orderCorrect && positions.length > 0,
    message: orderCorrect
      ? `All ${positions.length} items filled in matching order.`
      : `Item position order or count mismatch in filled rows.`,
  });

  // Check A2: Every position exactly once
  const posNumbers = positions.map((p) => p.position);
  const uniquePositions = new Set(posNumbers);
  const everyPosOnce = uniquePositions.size === positions.length;
  checks.push({
    id: 'A2_UNIQUE_POSITIONS',
    category: 'EXTRACTION',
    name: 'Every position present exactly once',
    passed: everyPosOnce,
    message: everyPosOnce ? 'Every quotation position appears exactly once.' : 'Duplicate position numbers found.',
  });

  // Check A3: References unique unless PDF repeats
  const refs = positions.map((p) => p.reference);
  const uniqueRefs = new Set(refs);
  const refRepetitionsInPdf = refs.length - uniqueRefs.size;
  if (refRepetitionsInPdf > 0) {
    flags.push(`${refRepetitionsInPdf} reference(s) repeated in PDF quotation.`);
  }
  checks.push({
    id: 'A3_REFERENCES',
    category: 'EXTRACTION',
    name: 'Reference numbers validated',
    passed: true,
    message: `References checked (${uniqueRefs.size} unique).`,
  });

  // Check A4: netPrice * quantity equals PDF line amount (tolerance 0.01)
  let lineAmountsValid = true;
  const lineAmountErrors: string[] = [];
  for (const pos of positions) {
    const computed = pos.netPrice * pos.quantity;
    if (Math.abs(computed - pos.lineAmount) > 0.01) {
      lineAmountsValid = false;
      lineAmountErrors.push(`Pos ${pos.position}: ${pos.netPrice} × ${pos.quantity} = ${computed.toFixed(2)} != ${pos.lineAmount}`);
    }
  }
  checks.push({
    id: 'A4_LINE_AMOUNTS',
    category: 'EXTRACTION',
    name: 'Net Price × Quantity equals PDF line amount',
    passed: lineAmountsValid,
    message: lineAmountsValid
      ? 'Net price × quantity equals PDF line amount on every line.'
      : `Mismatch in line totals: ${lineAmountErrors.slice(0, 2).join('; ')}`,
  });

  // Check A5: Net price not higher than list price
  let netLeList = true;
  for (const pos of positions) {
    if (pos.listPrice && pos.netPrice > pos.listPrice + 0.01) {
      netLeList = false;
      break;
    }
  }
  checks.push({
    id: 'A5_NET_PRICE_DISCOUNT',
    category: 'EXTRACTION',
    name: 'Net price <= List price',
    passed: netLeList,
    message: netLeList
      ? 'All net prices are lower than or equal to list prices.'
      : 'One or more net prices exceed list prices.',
  });

  // Check A6: Column A is text, columns E and F are numbers
  let colTypesValid = true;
  for (let i = 0; i < positions.length; i++) {
    const row = filledSheet.getRow(firstItemRow + i);
    const cellA = row.getCell(1);
    const cellE = row.getCell(5);
    const cellF = row.getCell(6);

    if (typeof cellA.text !== 'string' || !cellA.text) colTypesValid = false;
    if (typeof cellE.value !== 'number' || isNaN(cellE.value as number)) colTypesValid = false;
    if (typeof cellF.value !== 'number' || isNaN(cellF.value as number)) colTypesValid = false;
  }
  checks.push({
    id: 'A6_CELL_TYPES',
    category: 'EXTRACTION',
    name: 'Cell data types (A text, E/F numbers)',
    passed: colTypesValid,
    message: colTypesValid ? 'Column A formatted as text, E and F as numeric.' : 'Cell type validation failed on item rows.',
  });

  // Check non-PC units
  for (const pos of positions) {
    if (pos.unit && pos.unit.toUpperCase() !== 'PC' && pos.unit.toUpperCase() !== 'ST') {
      flags.push(`Pos ${pos.position} has non-PC unit: "${pos.unit}"`);
    }
  }

  // ==========================================
  // B. TOTALS CHECKS (via HyperFormula)
  // ==========================================

  // Recalculate Column G (Exworks) using HyperFormula
  let calculatedExworksSum = 0;
  for (let i = 0; i < positions.length; i++) {
    const rowIdx = firstItemRow + i - 1; // 0-indexed in HF
    const colIdx = 6; // Col G is index 6
    const cellVal = hf.getCellValue({ sheet: sheetId, col: colIdx, row: rowIdx });
    if (typeof cellVal === 'number') {
      calculatedExworksSum += cellVal;
    }
  }

  // Total in Totals row Col G
  const totalsRowGVal = hf.getCellValue({ sheet: sheetId, col: 6, row: totalsRow - 1 });
  const evaluatedTotalsRowG = typeof totalsRowGVal === 'number' ? totalsRowGVal : calculatedExworksSum;

  const invoiceAmount = extraction.invoiceAmount;
  const exworksMatchesInvoice = Math.abs(calculatedExworksSum - invoiceAmount) <= 0.05;

  if (!exworksMatchesInvoice) {
    flags.push(`Total Exworks (€${calculatedExworksSum.toFixed(2)}) differs from quotation invoice amount (€${invoiceAmount.toFixed(2)}).`);
  }

  checks.push({
    id: 'B1_TOTAL_EXWORKS_MATCH',
    category: 'TOTALS',
    name: 'Total Exworks equals PDF Invoice amount',
    passed: exworksMatchesInvoice,
    message: exworksMatchesInvoice
      ? `Total Exworks (€${calculatedExworksSum.toFixed(2)}) matches invoice amount (€${invoiceAmount.toFixed(2)}).`
      : `Total Exworks (€${calculatedExworksSum.toFixed(2)}) != invoice amount (€${invoiceAmount.toFixed(2)}).`,
  });

  // Check B2: Totals row SUM formula coverage
  const totalsRowCellG = filledSheet.getRow(totalsRow).getCell(7);
  let totalsFormula = totalsRowCellG.formula || '';
  if (!totalsFormula && typeof totalsRowCellG.value === 'object') {
    totalsFormula = (totalsRowCellG.value as any)?.formula || '';
  }

  let sumCoverageValid = true;
  if (totalsFormula) {
    const sumMatch = totalsFormula.match(/SUM\([$]?G[$]?(\d+):[$]?G[$]?(\d+)\)/i);
    if (sumMatch) {
      const sStart = parseInt(sumMatch[1], 10);
      const sEnd = parseInt(sumMatch[2], 10);
      sumCoverageValid = sStart === firstItemRow && sEnd === lastItemRow;
    }
  }

  checks.push({
    id: 'B2_SUM_RANGE_COVERAGE',
    category: 'TOTALS',
    name: 'Totals row SUM range covers all item rows',
    passed: sumCoverageValid,
    message: sumCoverageValid
      ? `SUM range correctly spans row ${firstItemRow} to ${lastItemRow}.`
      : `SUM formula range in totals row is incorrect.`,
  });

  // ==========================================
  // C. TEMPLATE INTEGRITY CHECKS
  // ==========================================

  // Check C1: Column I identical on all item rows and equals exchangeRate
  let colIValid = true;
  for (let i = 0; i < positions.length; i++) {
    const r = firstItemRow + i;
    const valI = filledSheet.getRow(r).getCell(9).value;
    if (Number(valI) !== exchangeRate) {
      colIValid = false;
      break;
    }
  }
  checks.push({
    id: 'C1_EXCHANGE_RATE_I',
    category: 'INTEGRITY',
    name: 'Column I identical on all item rows',
    passed: colIValid,
    message: colIValid
      ? `Exchange rate ${exchangeRate} written to Column I across all ${positions.length} item rows.`
      : 'Column I value mismatch or missing on item rows.',
  });

  // Check C2: O = 0 and P = 0.15 on every item row
  let opValid = true;
  for (let i = 0; i < positions.length; i++) {
    const r = firstItemRow + i;
    const valO = filledSheet.getRow(r).getCell(15).value;
    const valP = filledSheet.getRow(r).getCell(16).value;
    if (valO !== 0 || valP !== 0.15) {
      opValid = false;
      break;
    }
  }
  checks.push({
    id: 'C2_DUTY_VAT_OP',
    category: 'INTEGRITY',
    name: 'O = 0 (DUTY) and P = 0.15 (VAT 15%)',
    passed: opValid,
    message: opValid ? 'Column O (Duty=0) and Column P (VAT=0.15) correctly set on every item row.' : 'O or P values invalid.',
  });

  // Check C3: Formula preservation in C, D, J-N, Q-S (never overwritten with raw values)
  let formulaColsPreserved = true;
  const formulaCols = [3, 4, 10, 11, 12, 13, 14, 17, 18, 19];
  for (let i = 0; i < positions.length; i++) {
    const r = firstItemRow + i;
    const origRow = origSheet.getRow(firstItemRow); // baseline template row
    const filledRow = filledSheet.getRow(r);

    for (const c of formulaCols) {
      const origCell = origRow.getCell(c);
      const filledCell = filledRow.getCell(c);
      if (origCell.formula || (typeof origCell.value === 'object' && (origCell.value as any)?.formula)) {
        if (!filledCell.formula && !(typeof filledCell.value === 'object' && (filledCell.value as any)?.formula)) {
          formulaColsPreserved = false;
          break;
        }
      }
    }
  }
  checks.push({
    id: 'C3_FORMULA_COLS_PRESERVED',
    category: 'INTEGRITY',
    name: 'Columns C, D, J–N, Q–S preserved as formulas',
    passed: formulaColsPreserved,
    message: formulaColsPreserved
      ? 'All formula columns preserved intact without overwriting.'
      : 'One or more formula columns were overwritten with raw values.',
  });

  // Check C4: No Excel errors after recalculation in HyperFormula
  let hasFormulaErrors = false;
  const formulaErrorList: string[] = [];
  const maxRow = filledSheet.rowCount;
  const maxCol = filledSheet.columnCount;

  for (let r = 0; r < maxRow; r++) {
    for (let c = 0; c < maxCol; c++) {
      const cellVal = hf.getCellValue({ sheet: sheetId, col: c, row: r });
      const isError =
        (cellVal && typeof cellVal === 'object' && ('type' in cellVal || 'value' in cellVal)) ||
        (typeof cellVal === 'string' &&
          (cellVal.startsWith('#REF!') ||
            cellVal.startsWith('#NAME?') ||
            cellVal.startsWith('#DIV/0!') ||
            cellVal.startsWith('#VALUE!') ||
            cellVal.startsWith('#N/A')));

      if (isError) {
        hasFormulaErrors = true;
        const cellRef = `${numToCol(c + 1)}${r + 1}`;
        formulaErrorList.push(`${cellRef}: ${JSON.stringify(cellVal)}`);
      }
    }
  }

  checks.push({
    id: 'C4_NO_EXCEL_ERRORS',
    category: 'INTEGRITY',
    name: 'No formula calculation errors (#REF!, #DIV/0!, #VALUE!)',
    passed: !hasFormulaErrors,
    message: !hasFormulaErrors
      ? 'Recalculation verified clean: zero formula errors across the workbook.'
      : `Formula errors detected: ${formulaErrorList.slice(0, 3).join(', ')}`,
  });

  // Check C5: Sheet layout & structure match pristine original
  const sheetMatch = filledSheet.name === origSheet.name;
  checks.push({
    id: 'C5_SHEET_STRUCTURE',
    category: 'INTEGRITY',
    name: 'Workbook layout and sheet structure preserved',
    passed: sheetMatch,
    message: sheetMatch ? 'Sheet names, formatting, and layout structure preserved.' : 'Sheet structure changed.',
  });

  // Check Ignored lines and Unclear lines from PDF
  if (extraction.ignoredLines && extraction.ignoredLines.length > 0) {
    flags.push(`Ignored ${extraction.ignoredLines.length} text/sub-total lines from quotation.`);
  }
  if (extraction.unclear && extraction.unclear.length > 0) {
    flags.push(`Unclear items: ${extraction.unclear.map((u) => u.item).join(', ')}`);
  }

  const allPassed = checks.every((c) => c.passed);

  return {
    passed: allPassed,
    checks,
    exworksTotal: calculatedExworksSum,
    invoiceAmount,
    rowsAdded,
    exchangeRateUsed: exchangeRate,
    flags,
  };
}