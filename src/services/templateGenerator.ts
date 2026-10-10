import ExcelJS from 'exceljs';

/**
 * Generates a standard MULTIVAC Order List template (.xlsx) with proper styles and formulas
 */
export async function generateSampleTemplate(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MULTIVAC Systems';
  workbook.lastModifiedBy = 'MULTIVAC Systems';
  workbook.created = new Date();
  workbook.modified = new Date();

  const sheet = workbook.addWorksheet('MULTIVAC Order List', {
    views: [{ state: 'frozen', ySplit: 6 }],
  });

  // Set column widths
  sheet.columns = [
    { key: 'ref', width: 16 }, // A: Ref
    { key: 'name', width: 34 }, // B: Part Name
    { key: 'cur', width: 10 }, // C: Currency (Formula)
    { key: 'unit', width: 8 }, // D: Unit (Formula)
    { key: 'price', width: 14 }, // E: Net Price EUR
    { key: 'qty', width: 10 }, // F: Quantity
    { key: 'exworks', width: 16 }, // G: Total Exworks EUR (Formula)
    { key: 'rem', width: 14 }, // H: Remarks
    { key: 'rate', width: 14 }, // I: Exchange Rate
    { key: 'exmur', width: 16 }, // J: Exworks MUR (Formula)
    { key: 'freight', width: 14 }, // K: Freight (Formula)
    { key: 'cif', width: 16 }, // L: CIF MUR (Formula)
    { key: 'clear', width: 14 }, // M: Clearance (Formula)
    { key: 'landed', width: 16 }, // N: Landed Cost (Formula)
    { key: 'duty', width: 10 }, // O: Duty (0)
    { key: 'vat', width: 10 }, // P: VAT (0.15)
    { key: 'dutyMur', width: 14 }, // Q: Duty MUR (Formula)
    { key: 'vatMur', width: 14 }, // R: VAT MUR (Formula)
    { key: 'totalMur', width: 18 }, // S: Total Cost MUR (Formula)
  ];

  // Header rows
  sheet.mergeCells('A1:S1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = 'MULTIVAC SPARE PARTS ORDER LIST & COSTING';
  titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getRow(1).height = 32;

  sheet.getCell('A3').value = 'Customer:';
  sheet.getCell('B3').value = 'MULTIVAC Client Mauritius';
  sheet.getCell('A4').value = 'Quotation Ref:';
  sheet.getCell('B4').value = 'Q-2026-MV-0941';

  // Table Headers at Row 6
  const headers = [
    'Reference',
    'Part Name / Description',
    'Cur',
    'Unit',
    'Net Price (€)',
    'Qty',
    'Total Exworks (€)',
    'Remarks',
    'Exch Rate',
    'Exworks (MUR)',
    'Freight (MUR)',
    'CIF (MUR)',
    'Clearing (MUR)',
    'Landed (MUR)',
    'Duty',
    'VAT',
    'Duty (MUR)',
    'VAT (MUR)',
    'Total Cost (MUR)',
  ];

  const headerRow = sheet.getRow(6);
  headerRow.height = 24;
  headers.forEach((h, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = h;
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'medium', color: { argb: 'FF94A3B8' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    };
  });

  // Item rows: Row 7 to 16 (10 item rows)
  for (let r = 7; r <= 16; r++) {
    const row = sheet.getRow(r);
    row.height = 20;

    // Formulas
    row.getCell(3).value = { formula: `IF(A${r}="","","EUR")` }; // C
    row.getCell(4).value = { formula: `IF(A${r}="","","PC")` }; // D
    row.getCell(7).value = { formula: `IF(E${r}="","",E${r}*F${r})` }; // G: NetPrice * Qty
    row.getCell(10).value = { formula: `IF(G${r}="","",G${r}*I${r})` }; // J: Exworks MUR
    row.getCell(11).value = { formula: `IF(J${r}="","",J${r}*0.05)` }; // K: Freight 5%
    row.getCell(12).value = { formula: `IF(J${r}="","",J${r}+K${r})` }; // L: CIF MUR
    row.getCell(13).value = { formula: `IF(L${r}="","",L${r}*0.02)` }; // M: Clearing 2%
    row.getCell(14).value = { formula: `IF(L${r}="","",L${r}+M${r})` }; // N: Landed MUR
    row.getCell(17).value = { formula: `IF(N${r}="","",N${r}*O${r})` }; // Q: Duty MUR
    row.getCell(18).value = { formula: `IF(N${r}="","",N${r}*P${r})` }; // R: VAT MUR
    row.getCell(19).value = { formula: `IF(N${r}="","",N${r}+Q${r}+R${r})` }; // S: Total MUR

    // Cell styles & borders
    for (let c = 1; c <= 19; c++) {
      const cell = row.getCell(c);
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
      cell.font = { name: 'Arial', size: 10 };
      if (c === 5 || c === 7) {
        cell.numFmt = '#,##0.00';
      } else if (c === 6) {
        cell.numFmt = '#,##0';
      } else if (c === 9) {
        cell.numFmt = '0.00';
      } else if (c === 15 || c === 16) {
        cell.numFmt = '0.00%';
      } else if (c >= 10 && c <= 19) {
        cell.numFmt = '#,##0.00';
      }
    }
  }

  // Row 17: TOTALS ROW
  const totalsRow = sheet.getRow(17);
  totalsRow.height = 22;
  sheet.mergeCells('A17:D17');
  const totalLabel = sheet.getCell('A17');
  totalLabel.value = 'TOTAL EXWORKS';
  totalLabel.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F172A' } };
  totalLabel.alignment = { horizontal: 'right', vertical: 'middle' };

  totalsRow.getCell(6).value = { formula: 'SUM(F7:F16)' }; // Total Qty
  totalsRow.getCell(7).value = { formula: 'SUM(G7:G16)' }; // Total Exworks EUR
  totalsRow.getCell(10).value = { formula: 'SUM(J7:J16)' };
  totalsRow.getCell(19).value = { formula: 'SUM(S7:S16)' };

  for (let c = 1; c <= 19; c++) {
    const cell = totalsRow.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'double', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    };
    cell.font = { name: 'Arial', size: 10, bold: true };
    if (c === 7 || c === 10 || c === 19) {
      cell.numFmt = '#,##0.00';
    }
  }

  // Row 18: Cost and selling price block below
  const r18 = sheet.getRow(18);
  sheet.mergeCells('A18:D18');
  r18.getCell('A').value = 'Markup (25%)';
  r18.getCell(7).value = { formula: 'G17*0.25' };
  r18.getCell(7).numFmt = '#,##0.00';

  const r19 = sheet.getRow(19);
  sheet.mergeCells('A19:D19');
  r19.getCell('A').value = 'SELLING PRICE EXWORKS';
  r19.getCell('A').font = { bold: true };
  r19.getCell(7).value = { formula: 'G17+G18' };
  r19.getCell(7).font = { bold: true };
  r19.getCell(7).numFmt = '#,##0.00';

  const raw = await workbook.xlsx.writeBuffer();
  if (raw instanceof ArrayBuffer) {
    return raw;
  }
  const u8 = new Uint8Array(raw as any);
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
}
