/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import { Upload, FileSpreadsheet, FileText, CheckCircle2, AlertTriangle, XCircle, ArrowRight, Download, RefreshCw } from 'lucide-react';
import { fillTemplate, verifyOrderList, type VerificationReport } from './services/excelEngine';
import { generateSampleTemplate } from './services/templateGenerator';
import { generateSampleQuotationPdf } from './services/sampleQuotation';
import type { QuotationExtraction, PositionItem } from '../server';

type StepState = 'idle' | 'running' | 'finished' | 'uncertain';

export default function App() {
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [xlsxFile, setXlsxFile] = useState<File | null>(null);
  const [exchangeRate, setExchangeRate] = useState<string>('');
  const [exchangeRateTouched, setExchangeRateTouched] = useState<boolean>(false);

  const [step, setStep] = useState<StepState>('idle');
  const [progressMessage, setProgressMessage] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [uncertainties, setUncertainties] = useState<string[]>([]);

  // Results
  const [positions, setPositions] = useState<PositionItem[]>([]);
  const [verification, setVerification] = useState<VerificationReport | null>(null);
  const [filledExcelBuffer, setFilledExcelBuffer] = useState<ArrayBuffer | null>(null);

  const pdfInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);

  // Exchange rate validation
  const parsedRate = parseFloat(exchangeRate.trim());
  const isRateValid =
    exchangeRate.trim() !== '' &&
    !isNaN(parsedRate) &&
    parsedRate > 0 &&
    /^\d+(\.\d+)?$/.test(exchangeRate.trim());

  const canGenerate = pdfFile !== null && xlsxFile !== null && isRateValid && step !== 'running';

  // Helper to load sample files for quick testing
  const handleLoadSampleFiles = async () => {
    try {
      setErrorMessage(null);
      // Sample PDF
      const pdfBytes = generateSampleQuotationPdf();
      const samplePdf = new File(
        [pdfBytes.buffer.slice(pdfBytes.byteOffset, pdfBytes.byteOffset + pdfBytes.byteLength) as ArrayBuffer],
        'MULTIVAC_Quotation_Q-2026-MV-0941.pdf',
        {
          type: 'application/pdf',
        }
      );
      setPdfFile(samplePdf);

      // Sample Template
      const templateBuffer = await generateSampleTemplate();
      const sampleXlsx = new File([templateBuffer], 'Order_List_Template.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      setXlsxFile(sampleXlsx);

      // Exchange rate is intentionally left for user to enter
      setExchangeRate('50.25');
      setExchangeRateTouched(true);
    } catch (err: any) {
      setErrorMessage('Could not load sample files: ' + err.message);
    }
  };

  const handleStartOver = () => {
    setPdfFile(null);
    setXlsxFile(null);
    setExchangeRate('');
    setExchangeRateTouched(false);
    setStep('idle');
    setProgressMessage('');
    setErrorMessage(null);
    setUncertainties([]);
    setPositions([]);
    setVerification(null);
    setFilledExcelBuffer(null);
  };

  const handleGenerate = async () => {
    if (!canGenerate || !pdfFile || !xlsxFile) return;

    setStep('running');
    setErrorMessage(null);
    setUncertainties([]);
    setProgressMessage('Reading quotation…');

    try {
      // 1. Read PDF to base64
      const pdfBase64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          const res = reader.result as string;
          resolve(res.includes(',') ? res.split(',')[1] : res);
        };
        reader.onerror = reject;
        reader.readAsDataURL(pdfFile);
      });

      // 2. Call backend extraction (Double extraction with Gemini)
      setProgressMessage('Reading quotation and running verification passes…');
      const response = await fetch('/api/extract', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ pdfBase64 }),
      });

      const data = await response.json();

      if (!response.ok || !data.success) {
        if (data.uncertainties && data.uncertainties.length > 0) {
          setUncertainties(data.uncertainties);
          setStep('uncertain');
          setProgressMessage('');
          return;
        }
        throw new Error(data.error || data.message || 'Quotation extraction failed.');
      }

      const extraction: QuotationExtraction = data.extraction;
      setPositions(extraction.positions);

      // 3. Pristine template buffer
      setProgressMessage('Filling template…');
      const pristineBuffer = await xlsxFile.arrayBuffer();

      // Agent 2: Fill template deterministically
      const fillResult = await fillTemplate(pristineBuffer.slice(0), extraction, parsedRate);

      // 4. Verification with HyperFormula
      setProgressMessage('Verifying…');
      const verificationReport = await verifyOrderList(
        pristineBuffer.slice(0),
        fillResult.filledBuffer.slice(0),
        extraction,
        parsedRate,
        fillResult.layout,
        fillResult.rowsAdded
      );

      setVerification(verificationReport);
      setFilledExcelBuffer(fillResult.filledBuffer);
      setStep('finished');
      setProgressMessage('');
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || 'An error occurred during order list generation.');
      setStep('idle');
      setProgressMessage('');
    }
  };

  const handleDownload = () => {
    if (!filledExcelBuffer || !xlsxFile) return;
    const originalName = xlsxFile.name.replace(/\.xlsx$/i, '');
    const fileName = `${originalName}_filled.xlsx`;

    const blob = new Blob([filledExcelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col justify-center items-center py-16 px-4 font-sans selection:bg-slate-200">
      <main className="w-full max-w-xl bg-white border border-slate-200 rounded-xl shadow-xs p-8 sm:p-10 space-y-8">
        {/* Header */}
        <header className="space-y-1">
          <h1 className="text-2xl font-semibold tracking-tight text-slate-900">
            MULTIVAC Order List
          </h1>
          <p className="text-sm text-slate-500">
            Fill and verify MULTIVAC spare parts order templates from quotation documents.
          </p>
        </header>

        {/* Error message */}
        {errorMessage && (
          <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-lg text-sm text-rose-800 flex items-start gap-2.5">
            <XCircle className="w-4 h-4 text-rose-600 mt-0.5 shrink-0" />
            <div className="leading-snug">{errorMessage}</div>
          </div>
        )}

        {/* Uncertainties stop state */}
        {step === 'uncertain' && (
          <div className="space-y-4">
            <div className="p-4 bg-amber-50 border border-amber-200 rounded-lg space-y-2">
              <div className="flex items-center gap-2 text-amber-900 font-medium text-sm">
                <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                Uncertain values detected on quotation PDF
              </div>
              <p className="text-xs text-amber-800">
                Independent extraction passes could not definitively confirm the following values without guessing. File generation stopped:
              </p>
              <ul className="text-xs text-amber-900 list-disc list-inside space-y-1 font-mono pt-1">
                {uncertainties.map((u, i) => (
                  <li key={i}>{u}</li>
                ))}
              </ul>
            </div>
            <button
              onClick={handleStartOver}
              className="text-xs text-slate-500 hover:text-slate-800 underline transition-colors cursor-pointer"
            >
              Start over
            </button>
          </div>
        )}

        {/* INPUT FORM (idle or running) */}
        {(step === 'idle' || step === 'running') && (
          <div className="space-y-6">
            {/* Input 1: Quotation PDF */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600">
                Quotation PDF (.pdf)
              </label>
              <div
                onClick={() => pdfInputRef.current?.click()}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  const file = e.dataTransfer.files[0];
                  if (file && (file.type === 'application/pdf' || file.name.endsWith('.pdf'))) {
                    setPdfFile(file);
                  }
                }}
                className={`border border-dashed rounded-lg p-4 cursor-pointer transition-colors flex items-center justify-between gap-3 ${
                  pdfFile
                    ? 'border-slate-400 bg-slate-50'
                    : 'border-slate-300 hover:border-slate-400 bg-white hover:bg-slate-50/50'
                }`}
              >
                <div className="flex items-center gap-3 overflow-hidden">
                  <div className="p-2 rounded-md bg-slate-100 text-slate-600 shrink-0">
                    <FileText className="w-5 h-5" />
                  </div>
                  <div className="truncate">
                    {pdfFile ? (
                      <p className="text-sm font-medium text-slate-900 truncate">{pdfFile.name}</p>
                    ) : (
                      <p className="text-sm text-slate-500">
                        Click or drag MULTIVAC quotation PDF here
                      </p>
                    )}
                    <p className="text-xs text-slate-400">
                      {pdfFile ? `${(pdfFile.size / 1024).toFixed(0)} KB` : 'PDF only'}
                    </p>
                  </div>
                </div>
                {pdfFile && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      setPdfFile(null);
                    }}
                    className="text-xs text-slate-400 hover:text-slate-600 px-2 py-1"
                  >
                    Change
                  </button>
                )}
              </div>
              <input
                ref={pdfInputRef}
                type="file"
                accept=".pdf,application/pdf"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) setPdfFile(file);
                }}
              />
            </div>

            {/* Input 2: Order List template */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600">
                Order List template (.xlsx)
              </label>
              <div
                onClick={() => xlsxInputRef.current?.click()}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  const file = e.dataTransfer.files[0];
                  if (file && (file.name.endsWith('.xlsx') || file.type.includes('spreadsheet'))) {
                    setXlsxFile(file);
                  }
                }}
                className={`border border-dashed rounded-lg p-4 cursor-pointer transition-colors flex items-center justify-between gap-3 ${
                  xlsxFile
                    ? 'border-slate-400 bg-slate-50'
                    : 'border-slate-300 hover:border-slate-400 bg-white hover:bg-slate-50/50'
                }`}
              >
                <div className="flex items-center gap-3 overflow-hidden">
                  <div className="p-2 rounded-md bg-slate-100 text-slate-600 shrink-0">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div className="truncate">
                    {xlsxFile ? (
                      <p className="text-sm font-medium text-slate-900 truncate">{xlsxFile.name}</p>
                    ) : (
                      <p className="text-sm text-slate-500">
                        Click or drag Excel order-list template here
                      </p>
                    )}
                    <p className="text-xs text-slate-400">
                      {xlsxFile ? `${(xlsxFile.size / 1024).toFixed(0)} KB` : 'Excel .xlsx only'}
                    </p>
                  </div>
                </div>
                {xlsxFile && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      setXlsxFile(null);
                    }}
                    className="text-xs text-slate-400 hover:text-slate-600 px-2 py-1"
                  >
                    Change
                  </button>
                )}
              </div>
              <input
                ref={xlsxInputRef}
                type="file"
                accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) setXlsxFile(file);
                }}
              />
            </div>

            {/* Input 3: Exchange rate */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold uppercase tracking-wider text-slate-600">
                EUR → MUR exchange rate
              </label>
              <div className="relative">
                <input
                  type="text"
                  inputMode="decimal"
                  placeholder="Enter positive number (e.g. 50.25)"
                  value={exchangeRate}
                  disabled={step === 'running'}
                  onChange={(e) => {
                    setExchangeRate(e.target.value);
                    setExchangeRateTouched(true);
                  }}
                  onBlur={() => setExchangeRateTouched(true)}
                  className={`w-full px-3.5 py-2.5 bg-white border rounded-lg text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-1 focus:ring-slate-900 transition-colors ${
                    exchangeRateTouched && !isRateValid
                      ? 'border-rose-400 focus:border-rose-500'
                      : 'border-slate-300'
                  }`}
                />
              </div>
              {/* Short inline message when invalid */}
              {exchangeRateTouched && !isRateValid && (
                <p className="text-xs text-rose-600">
                  {exchangeRate.trim() === ''
                    ? 'Please enter the EUR → MUR exchange rate.'
                    : 'Exchange rate must be a single positive number.'}
                </p>
              )}
            </div>

            {/* Sample files helper */}
            {!pdfFile && !xlsxFile && (
              <div className="pt-1">
                <button
                  type="button"
                  onClick={handleLoadSampleFiles}
                  className="text-xs text-slate-500 hover:text-slate-900 underline underline-offset-2 transition-colors cursor-pointer"
                >
                  Load sample MULTIVAC files for testing
                </button>
              </div>
            )}

            {/* Primary Action Button */}
            <div className="pt-2">
              <button
                type="button"
                disabled={!canGenerate}
                onClick={handleGenerate}
                className={`w-full py-2.5 px-4 rounded-lg text-sm font-medium transition-colors flex items-center justify-center gap-2 cursor-pointer ${
                  canGenerate
                    ? 'bg-slate-900 text-white hover:bg-slate-800 active:bg-slate-950 shadow-xs'
                    : 'bg-slate-100 text-slate-400 cursor-not-allowed border border-slate-200'
                }`}
              >
                {step === 'running' ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    <span>Processing…</span>
                  </>
                ) : (
                  <>
                    <span>Generate order list</span>
                    <ArrowRight className="w-4 h-4" />
                  </>
                )}
              </button>
            </div>

            {/* Simple progress line */}
            {step === 'running' && progressMessage && (
              <p className="text-xs text-slate-500 text-center animate-pulse">
                {progressMessage}
              </p>
            )}
          </div>
        )}

        {/* FINISHED STATE */}
        {step === 'finished' && verification && (
          <div className="space-y-6">
            {/* Status Badge & Check results */}
            <div className="space-y-3">
              {verification.passed ? (
                <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-semibold tracking-wide">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                  All checks passed
                </div>
              ) : (
                <div className="space-y-2">
                  <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-rose-50 border border-rose-200 text-rose-800 text-xs font-semibold tracking-wide">
                    <XCircle className="w-4 h-4 text-rose-600" />
                    Verification failed
                  </div>
                  <ul className="text-xs text-rose-700 space-y-1 list-disc list-inside bg-rose-50/70 p-3 rounded-lg border border-rose-200">
                    {verification.checks
                      .filter((c) => !c.passed)
                      .map((c) => (
                        <li key={c.id}>
                          <span className="font-medium">{c.name}:</span> {c.message}
                        </li>
                      ))}
                  </ul>
                </div>
              )}

              {/* Single Summary Line */}
              <p className="text-xs text-slate-600 leading-relaxed font-mono">
                Exchange rate used: {verification.exchangeRateUsed} · Rows added:{' '}
                {verification.rowsAdded} · Total Exworks: €
                {verification.exworksTotal.toFixed(2)} vs Invoice amount: €
                {verification.invoiceAmount.toFixed(2)}
              </p>
            </div>

            {/* Flags for unusual things */}
            {verification.flags.length > 0 && (
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-1.5">
                <p className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
                  Notices & Flags
                </p>
                <ul className="text-xs text-slate-700 space-y-1">
                  {verification.flags.map((flag, idx) => (
                    <li key={idx} className="flex items-start gap-1.5">
                      <span className="text-slate-400 mt-0.5">•</span>
                      <span>{flag}</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {/* Compact Table */}
            <div className="space-y-2">
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-600">
                Extracted Line Items ({positions.length})
              </p>
              <div className="border border-slate-200 rounded-lg overflow-hidden">
                <div className="overflow-x-auto max-h-56">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-100 text-slate-700 uppercase tracking-wider text-[10px] sticky top-0 border-b border-slate-200">
                      <tr>
                        <th className="py-2 px-3">Reference</th>
                        <th className="py-2 px-3">Part Name</th>
                        <th className="py-2 px-3 text-right">Net Price</th>
                        <th className="py-2 px-3 text-right">Qty</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-800">
                      {positions.map((item, idx) => (
                        <tr key={idx} className="hover:bg-slate-50">
                          <td className="py-2 px-3 font-mono font-medium text-slate-900">
                            {item.reference}
                          </td>
                          <td className="py-2 px-3 truncate max-w-[200px]" title={item.partName}>
                            {item.partName}
                          </td>
                          <td className="py-2 px-3 text-right font-mono">
                            €{item.netPrice.toFixed(2)}
                          </td>
                          <td className="py-2 px-3 text-right font-mono">
                            {item.quantity} {item.unit}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Download Button (available only if all checks passed) */}
            <div className="pt-2 space-y-3">
              {verification.passed ? (
                <button
                  type="button"
                  onClick={handleDownload}
                  className="w-full py-2.5 px-4 rounded-lg text-sm font-medium bg-emerald-600 hover:bg-emerald-700 text-white transition-colors flex items-center justify-center gap-2 shadow-xs cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>Download filled Excel</span>
                </button>
              ) : (
                <p className="text-xs text-rose-600 text-center">
                  Download unavailable until all verification checks pass.
                </p>
              )}

              {/* Small Start over link */}
              <div className="text-center">
                <button
                  type="button"
                  onClick={handleStartOver}
                  className="text-xs text-slate-400 hover:text-slate-700 underline underline-offset-2 transition-colors cursor-pointer"
                >
                  Start over
                </button>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
