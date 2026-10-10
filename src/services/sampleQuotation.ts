/**
 * Generates a valid PDF 1.4 representing a standard MULTIVAC quotation document
 */
export function generateSampleQuotationPdf(): Uint8Array {
  const contentStream = `BT
/F1 16 Tf
50 780 Td
(MULTIVAC Sepp Haggenmueller SE & Co. KG) Tj
/F1 10 Tf
0 -18 Td
(Bahnhofstr. 4, 87787 Wolfertschwenden, Germany) Tj
0 -25 Td
/F1 14 Tf
(QUOTATION / SPARE PARTS OFFER No. Q-2026-MV-0941) Tj
/F1 10 Tf
0 -20 Td
(Date: 15.03.2026   Customer No: 489210   Currency: EUR) Tj
0 -15 Td
(Delivery Terms: EXW Wolfertschwenden / Ex Works) Tj
0 -25 Td
/F1 10 Tf
(Pos.  Reference     Description                            Qty   Unit  List Price  Net Price   Amount EUR) Tj
0 -15 Td
(---------------------------------------------------------------------------------------------------------) Tj
0 -18 Td
(10    108342110     Tension spring D1=9 d=1.2 L0=45        2     PC    14,50       12,30       24,60) Tj
0 -18 Td
(20    092415782     Seal silicone profile 12x8             1     m     45,00       38,25       38,25) Tj
0 -18 Td
(30    104523991     Heating wire 3.5x0.2                   4     PC    8,20        6,97        27,88) Tj
0 -18 Td
(40    107612340     Cut-off blade serrated L=420           1     PC    115,00      97,75       97,75) Tj
0 -25 Td
(Sub-total spare parts net:                                                                 188,48) Tj
0 -15 Td
(Freight & Insurance: To be determined on dispatch) Tj
0 -20 Td
/F1 12 Tf
(Invoice amount:                                                                     EUR 188,48) Tj
0 -30 Td
/F1 9 Tf
(Note: Net price reflects applied customer contractual discount. Payment 30 days net.) Tj
ET`;

  const streamLength = contentStream.length;

  const pdf = `%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length ${streamLength} >>
stream
${contentStream}
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000227 00000 n 
0000000${(300 + streamLength).toString().padStart(3, '0')} 00000 n 
trailer
<< /Size 6 /Root 1 0 R >>
startxref
${400 + streamLength}
%%EOF`;

  return new TextEncoder().encode(pdf);
}
