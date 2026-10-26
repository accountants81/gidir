import { useEffect, useRef, useState } from 'react';
import { ArrowDownToLine, ArrowLeft, ArrowRight, Check, CheckCircle2, FileSpreadsheet, FolderOpen, Languages, LockKeyhole, Moon, RotateCcw, ShieldCheck, Sun, WifiOff, X, AlertTriangle } from 'lucide-react';
import { motion } from 'framer-motion';
import { Stepper } from '@/components/Stepper';
import { WorkbookPreview } from '@/components/WorkbookPreview';
import { ConversionReport } from '@/components/ConversionReport';
import { MissingValuesEditor } from '@/components/MissingValuesEditor';
import { DataCorrectionsEditor } from '@/components/DataCorrectionsEditor';
import { InvoiceFilters, type FilterSelection } from '@/components/InvoiceFilters';
import { generateTaxCSVParts, getWorkbookSheetNames, parseWorkbook, revalidateWorkbook, safeFilename } from '@/lib/local-converter';
import type { CsvEncoding } from '@/lib/local-converter';
import { HEADERS, isColumnRequired } from '@/lib/schemas';
import type { Direction, ParsedWorkbook, ValidationIssue } from '@/types/converter';
import './App.css';

type InstallPrompt = Event & { prompt: () => Promise<void>; userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }> };
type Theme = 'dark' | 'light';
type ExportKind = 'final' | 'draft';
type OutputFile = { filename: string; bytes: Uint8Array };
function App() {
  const [english, setEnglish] = useState(false);
  const [theme, setTheme] = useState<Theme>(() => {
    try {
      return window.localStorage.getItem('tax-converter-theme') === 'light' ? 'light' : 'dark';
    } catch {
      return 'dark';
    }
  });
  const [direction, setDirection] = useState<Direction | null>(null);
  const [step, setStep] = useState(1);
  const [workbook, setWorkbook] = useState<ParsedWorkbook | null>(null);
  const [filename, setFilename] = useState('');
  const [csvEncoding, setCsvEncoding] = useState<CsvEncoding>('windows-1256');
  const [lastDownloadedEncoding, setLastDownloadedEncoding] = useState<CsvEncoding>('windows-1256');
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState('');
  const [exportError, setExportError] = useState('');
  const [confirm, setConfirm] = useState(false);
  const [confirmMode, setConfirmMode] = useState<'draft' | 'warnings'>('warnings');
  const [exportKind, setExportKind] = useState<ExportKind>('final');
  const [outputFiles, setOutputFiles] = useState<OutputFile[]>([]);
  const [installEvent, setInstallEvent] = useState<InstallPrompt | null>(null);
  const [iosInstall, setIosInstall] = useState(false);
  const [showIosHelp, setShowIosHelp] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const t = (ar: string, en: string) => english ? en : ar;
  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    document.documentElement.style.colorScheme = theme;
    try {
      window.localStorage.setItem('tax-converter-theme', theme);
    } catch {
      // The theme still works for this session when storage is unavailable.
    }
  }, [theme]);
  useEffect(() => {
    const handler = (event: Event) => { event.preventDefault(); setInstallEvent(event as InstallPrompt); };
    window.addEventListener('beforeinstallprompt', handler);
    const userAgent = window.navigator.userAgent;
    setIosInstall(/iPad|iPhone|iPod/.test(userAgent) || (window.navigator.platform === 'MacIntel' && window.navigator.maxTouchPoints > 1));
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);
  const chooseDirection = (value: Direction) => {
    setDirection(value);
    setWorkbook(null);
    setPendingFile(null);
    setSheetNames([]);
    setSelectedSheet('');
    setError('');
    setExportError('');
    setOutputFiles([]);
    setExportKind('final');
    setStep(1);
  };
  const showReadError = (e: unknown) => {
    const message = e instanceof Error ? e.message : '';
    const arMessage = message.includes('Excel workbook') ? 'يرجى اختيار ملف Excel بصيغة XLSX أو XLS.' : message.includes('worksheet') ? 'لا يحتوي هذا الملف على ورقة عمل.' : message.includes('first row') ? 'يجب أن يحتوي الصف الأول على عناوين الأعمدة الأصلية.' : message.includes('tax template') ? 'عناوين الأعمدة لا تطابق نموذج المبيعات أو المشتريات.' : 'تعذرت قراءة الملف. تحقق من سلامة المصنف وحاول مرة أخرى.';
    setError(english ? (message || 'Could not read this workbook.') : arMessage);
  };
  const loadSelectedSheet = async (file: File, sheetName: string) => {
    if (!direction) return;
    setError('');
    setExportError('');
    setBusy(true);
    try {
      const parsed = await parseWorkbook(file, direction, sheetName);
      setWorkbook(parsed);
      const now = new Date();
      setFilename(`${parsed.direction}_final_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`);
      setPendingFile(null);
      setSheetNames([]);
      setSelectedSheet('');
      setStep(2);
    } catch (e) {
      showReadError(e);
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = '';
    }
  };
  const loadFile = async (file?: File) => {
    if (!file || !direction) return;
    setError('');
    setBusy(true);
    try {
      const names = await getWorkbookSheetNames(file);
      if (names.length > 1) {
        setPendingFile(file);
        setSheetNames(names);
        setSelectedSheet('');
        return;
      }
      const parsed = await parseWorkbook(file, direction, names[0]);
      setWorkbook(parsed);
      const now = new Date();
      setFilename(`${parsed.direction}_final_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`);
      setStep(2);
    } catch (e) {
      showReadError(e);
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = '';
    }
  };
  const downloadCsvPart = (file: OutputFile) => {
    const blob = new Blob([file.bytes.slice().buffer as ArrayBuffer], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = file.filename;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  const exportCsv = (kind: ExportKind = 'final', encoding: CsvEncoding = csvEncoding) => {
    if (!workbook || !direction) return;
    setExportError('');
    try {
      const parts = generateTaxCSVParts(workbook.headers, workbook.rows, encoding);
      const baseName = kind === 'draft'
        ? safeFilename(`${workbook.direction}_review_incomplete`, workbook.direction)
        : safeFilename(filename, workbook.direction);
      const stem = baseName.replace(/\.csv$/i, '');
      const files = parts.map((bytes, index) => ({
        filename: parts.length === 1 ? baseName : `${stem}_part_${index + 1}.csv`,
        bytes,
      }));
      setOutputFiles(files);
      setExportKind(kind);
      if (files[0]) downloadCsvPart(files[0]);
      setLastDownloadedEncoding(encoding);
      setConfirm(false);
      setConfirmMode('warnings');
      setStep(3);
    } catch (e) {
      setExportError(e instanceof Error ? e.message : (english ? 'Could not encode this CSV.' : 'تعذر ترميز ملف CSV.'));
    }
  };
  const requestExport = () => {
    if (!workbook) return;
    const missingRequired = workbook.rows.reduce((count, row) =>
      count + row.reduce((rowCount, _value, column) => rowCount + (isColumnRequired(workbook.direction, row, column) && !row[column]?.trim() ? 1 : 0), 0), 0);
    const hasErrors = workbook.issues.some((issue) => issue.severity === 'error');
    const hasWarnings = workbook.issues.some((issue) => issue.severity === 'warning');
    if (missingRequired > 0 || hasErrors) {
      setConfirmMode('draft');
      setConfirm(true);
    } else if (hasWarnings) {
      setConfirmMode('warnings');
      setConfirm(true);
    }
    else exportCsv();
  };
  const applyColumnValue = (column: number, value: string, rowIndexes: number[]) => {
    setWorkbook((current) => {
      if (!current) return current;
      const rows = current.rows.map((row, index) => {
        if (!rowIndexes.includes(index) || row[column]?.trim()) return row;
        const updated = [...row];
        updated[column] = value;
        return updated;
      });
      return revalidateWorkbook(current, rows);
    });
  };
  const applyCorrection = (issue: ValidationIssue, value: string) => {
    setWorkbook((current) => {
      if (!current) return current;
      const rowIndex = current.sourceRowNumbers?.indexOf(issue.rowNumber) ?? -1;
      if (rowIndex < 0) return current;
      const rows = current.rows.map((row, index) => {
        if (index !== rowIndex) return row;
        const updated = [...row];
        updated[issue.columnIndex] = value;
        return updated;
      });
      return revalidateWorkbook(current, rows);
    });
  };
  const removeRows = (rowIndexes: number[], reasonAr: string, reason: string) => {
    setWorkbook((current) => {
      if (!current) return current;
      const removedIndexes = new Set(rowIndexes);
      const keptIndexes = current.rows.map((_row, index) => index).filter((index) => !removedIndexes.has(index));
      const rows = keptIndexes.map((index) => current.rows[index]);
      const sourceRowNumbers = keptIndexes.map((index) => current.sourceRowNumbers?.[index] ?? index + 2);
      const filtered = revalidateWorkbook({ ...current, sourceRowNumbers }, rows);
      return {
        ...filtered,
        normalizations: [...filtered.normalizations, {
          code: `explicit-row-removal-${Date.now()}`,
          count: rowIndexes.length,
          reason,
          reasonAr,
        }],
      };
    });
  };
  const applyInvoiceFilter = (selection: FilterSelection, keptIndexes: number[]) => {
    setWorkbook((current) => {
      if (!current) return current;
      const rows = keptIndexes.map((index) => current.rows[index]);
      const sourceRowNumbers = keptIndexes.map((index) => current.sourceRowNumbers?.[index] ?? index + 2);
      const filtered = revalidateWorkbook({ ...current, sourceRowNumbers }, rows);
      const removed = current.rows.length - rows.length;
      const period = selection.month && selection.year ? `${selection.month}/${selection.year}` : selection.month ? `الشهر ${selection.month}` : selection.year ? `السنة ${selection.year}` : '';
      const reasonAr = `تم حذف ${removed} صفوف بناءً على تصفية اختارها المستخدم${period ? ` (${period})` : ''}${selection.removeCancelled ? ' وحذف نوع المستند 3' : ''}.`;
      return { ...filtered, normalizations: [...filtered.normalizations, { code: `invoice-filter-${Date.now()}`, count: removed, reason: `Removed ${removed} rows using the user's optional invoice filter.`, reasonAr }] };
    });
  };
  const applySingleItemDefaults = (rowIndexes: number[]) => {
    setWorkbook((current) => {
      if (!current) return current;
      const rows = current.rows.map((row, index) => {
        if (!rowIndexes.includes(index)) return row;
        const updated = [...row];
        if (!updated[18]?.trim()) updated[18] = '1';
        if (!updated[16]?.trim() && updated[19]?.trim()) updated[16] = updated[19];
        return updated;
      });
      return revalidateWorkbook(current, rows);
    });
  };
  const startOver = () => {
    setWorkbook(null);
    setDirection(null);
    setPendingFile(null);
    setSheetNames([]);
    setSelectedSheet('');
    setStep(1);
    setFilename('');
    setError('');
    setExportError('');
    setOutputFiles([]);
    setExportKind('final');
  };
  const install = async () => {
    if (iosInstall) {
      setShowIosHelp(true);
      return;
    }
    if (!installEvent) return;
    await installEvent.prompt();
    await installEvent.userChoice;
    setInstallEvent(null);
  };
  const headings = [
    t('اختَر نوع الفاتورة', 'Choose invoice type'),
    t('راجع بياناتك', 'Review your data'),
    exportKind === 'draft'
      ? t('مسودة مراجعة غير جاهزة للرفع', 'Review draft — not ready to upload')
      : t('ملفك جاهز', 'Your file is ready'),
  ];
  return <main className={`app-shell ${english ? 'is-english' : ''}`} dir={english ? 'ltr' : 'rtl'} data-theme={theme}>
    <header className="topbar">
      <a className="brand" href={import.meta.env.BASE_URL} aria-label={t('الرئيسية', 'Home')}>
        <span className="brand-mark" aria-hidden="true">ض</span>
        <span className="brand-copy"><strong>{t('محول فواتير الضرائب', 'Tax Invoice Converter')}</strong><small>{t('تحويل محلي وآمن', 'Private, on-device conversion')}</small></span>
      </a>
      <div className="top-actions">
        <span className="trust-badge"><WifiOff size={13} />{t('يعمل دون اتصال', 'Offline ready')}</span>
        <span className="trust-badge privacy"><LockKeyhole size={13} />{t('خصوصيتك محفوظة', 'Private by design')}</span>
        {(installEvent || iosInstall) && <button className="install-link" onClick={install}>{t('تثبيت التطبيق', 'Install app')}</button>}
        <button className="theme-toggle" type="button" onClick={() => setTheme((value) => value === 'dark' ? 'light' : 'dark')} aria-label={theme === 'dark' ? t('التبديل إلى وضع النهار', 'Switch to light mode') : t('التبديل إلى وضع الليل', 'Switch to dark mode')} aria-pressed={theme === 'light'} data-testid="button-toggle-theme">
          {theme === 'dark' ? <Sun size={15} aria-hidden="true" /> : <Moon size={15} aria-hidden="true" />}
          <span>{theme === 'dark' ? t('نهار', 'Light') : t('ليل', 'Dark')}</span>
        </button>
        <button className="language-toggle" onClick={() => setEnglish((value) => !value)} aria-label={t('Switch language to English', 'التبديل إلى العربية')}><Languages size={16} /><span>{english ? 'العربية' : 'English'}</span></button>
      </div>
    </header>
    <motion.section className="hero" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.45 }}>
      <div className="hero-orbit orbit-one" /><div className="hero-orbit orbit-two" />
      <div className="hero-content fade-up">
          <div className="compatibility"><span className="compat-dot" />{t('مهيأ لتنسيق ملفات الضرائب المصرية', 'Structured for Egyptian tax workflows')}</div>
          <h1>{t('حوّل فواتيرك بسهولة،', 'Convert your tax invoices,')}<br /><em>{t('من Excel إلى CSV.', 'from Excel to CSV.')}</em></h1>
          <p>{t('راجع بياناتك أولاً، ثم نزّل ملف CSV. تتم المعالجة على جهازك دون رفع الملفات.', 'Review your data, then download a CSV. Your workbook is processed on your device, never uploaded.')}</p>
      </div>
      <div className="trust-strip"><span><ShieldCheck size={15} />{t('معالجة محلية بالكامل', 'Processed entirely on-device')}</span><i /><span><FileSpreadsheet size={15} />{t('Excel إلى CSV', 'Excel to CSV')}</span><i /><span>{t('لا تسجيل دخول مطلوب', 'No account needed')}</span></div>
    </motion.section>
    <section className="workspace">
      <div className="flow-heading"><div><span className="eyebrow">{t('مساحة التحويل', 'CONVERSION WORKSPACE')}</span><h2>{headings[step === 3 ? 2 : workbook ? 1 : 0]}</h2></div><Stepper step={step} labels={[t('النوع', 'Type'), t('المراجعة', 'Review'), t('التنزيل', 'Download')]} /></div>
      {!workbook && <motion.section className="glass selection-panel fade-up" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.4 }}>
        <div className="section-intro"><span className="section-index">01</span><div><h3>{t('ما نوع المعاملة؟', 'What are you converting?')}</h3><p>{t('اختر نوع الملف لتبدأ. يمكنك تغييره لاحقاً.', 'Choose a workbook type to get started. You can switch later.')}</p></div></div>
        <div className="type-options">
          <button className={`type-option ${direction === 'sales' ? 'selected' : ''}`} onClick={() => chooseDirection('sales')} aria-pressed={direction === 'sales'} data-testid="button-select-sales">
            <span className="type-icon"><ArrowDownToLine size={20} /></span><span className="type-label"><b>{t('فواتير المبيعات', 'Sales invoices')}</b><small>{t('إقرارات المبيعات المحلية', 'Domestic sales records')}</small></span><span className="radio-dot">{direction === 'sales' && <Check size={12} />}</span>
          </button>
          <button className={`type-option ${direction === 'purchases' ? 'selected' : ''}`} onClick={() => chooseDirection('purchases')} aria-pressed={direction === 'purchases'} data-testid="button-select-purchases">
            <span className="type-icon purchase"><ArrowDownToLine size={20} /></span><span className="type-label"><b>{t('فواتير المشتريات', 'Purchase invoices')}</b><small>{t('إقرارات المشتريات والموردين', 'Purchases and supplier records')}</small></span><span className="radio-dot">{direction === 'purchases' && <Check size={12} />}</span>
          </button>
        </div>
        {direction && <div className="upload-zone-wrap fade-up">
          <input ref={inputRef} id="workbook-file" className="visually-hidden" type="file" accept=".xlsx,.xls,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={(event) => void loadFile(event.target.files?.[0])} aria-label={t('اختيار ملف Excel', 'Choose Excel workbook')} />
          <div className={`upload-zone ${dragging ? 'dragging' : ''}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); void loadFile(event.dataTransfer.files[0]); }}>
            <span className="folder-float float-folder"><FolderOpen size={27} /></span><div className="upload-copy"><b>{busy ? t('جارٍ قراءة الملف محلياً…', 'Reading workbook locally…') : t('أفلت ملف Excel هنا', 'Drop your Excel workbook here')}</b><span>{t('أو اختره من جهازك · XLSX أو XLS', 'or browse your device · XLSX or XLS')}</span></div>
            <button className="browse-button" onClick={() => inputRef.current?.click()} disabled={busy}>{t('اختيار ملف', 'Browse files')}</button>
          </div>
          {error && <div className="error-banner" role="alert"><AlertTriangle size={17} /><span>{error}</span><button onClick={() => setError('')} aria-label={t('إغلاق', 'Dismiss')}><X size={15} /></button></div>}
          <p className="upload-assurance"><LockKeyhole size={13} />{t('يتم فتح الملف داخل المتصفح فقط — لا رفع، لا تخزين.', 'Your workbook stays in this browser — never uploaded or stored.')}</p>
          {pendingFile && sheetNames.length > 1 && <div className="sheet-picker">
            <label htmlFor="workbook-sheet">{t('الملف يحتوي على أكثر من ورقة. اختر الورقة المطلوبة:', 'This workbook has multiple sheets. Choose the sheet to use:')}</label>
            <div>
              <select id="workbook-sheet" value={selectedSheet} onChange={(event) => setSelectedSheet(event.target.value)}>
                <option value="">{t('اختر ورقة العمل', 'Choose a worksheet')}</option>
                {sheetNames.map((name) => <option key={name} value={name}>{name}</option>)}
              </select>
              <button
                type="button"
                className="gold-button"
                disabled={!selectedSheet || busy}
                onClick={() => { if (pendingFile && selectedSheet) void loadSelectedSheet(pendingFile, selectedSheet); }}
              >
                {busy ? t('جارٍ القراءة…', 'Reading…') : t('متابعة بهذه الورقة', 'Continue with this sheet')}
              </button>
            </div>
          </div>}
          <details className="schema-help"><summary>{t('عرض عناوين المبيعات والمشتريات المقبولة', 'View accepted sales and purchase headers')}</summary>{(['sales', 'purchases'] as const).map((templateDirection) => <div key={templateDirection}><p className="schema-template-label">{templateDirection === 'sales' ? t('المبيعات', 'Sales') : t('المشتريات', 'Purchases')}</p><ol>{HEADERS[templateDirection].map((header, index) => <li key={index}><span>{index + 1}</span><b dir="auto">{header}</b></li>)}</ol></div>)}</details>
        </div>}
      </motion.section>}
      {workbook && step === 2 && <div className="review-layout fade-up">
        <div className="review-toolbar"><div className="section-intro compact"><span className="section-index">02</span><div><h3>{t('معاينة وتحقق', 'Preview & validation')}</h3><p>{t('راجع الصفوف والتنبيهات قبل إنشاء ملف CSV.', 'Review rows and findings before creating your CSV.')}</p></div></div><button className="text-button" onClick={() => { setWorkbook(null); setStep(1); setOutputFiles([]); }}><RotateCcw size={14} />{t('استبدال الملف', 'Replace file')}</button></div>
        <MissingValuesEditor workbook={workbook} english={english} onApply={applyColumnValue} onApplySingleItemDefaults={applySingleItemDefaults} />
        <DataCorrectionsEditor workbook={workbook} english={english} onApply={applyCorrection} onRemoveRows={removeRows} />
        <WorkbookPreview workbook={workbook} english={english} />
        <InvoiceFilters workbook={workbook} english={english} onApply={applyInvoiceFilter} />
        <ConversionReport workbook={workbook} encoding={csvEncoding} english={english} />
        <section className="glass export-panel">
          <div className="export-copy"><span className="section-index">03</span><div>
            <h3>{t('إعداد ملف التصدير', 'Export settings')}</h3>
            <p>{csvEncoding === 'windows-1256'
              ? t('الافتراضي: CSV قديم بترميز Windows-1256 دون BOM، مناسب لبوابة الضرائب القديمة.', 'Default: legacy CSV in Windows-1256 without BOM for the older tax portal.')
              : t('نسخة UTF-8 مع BOM للتطبيقات الحديثة.', 'UTF-8 with BOM for modern spreadsheet apps.')}</p>
          </div></div>
          <label className="filename-control"><span>{t('اسم الملف', 'Filename')}</span><div className="filename-input"><input value={filename} onChange={(event) => setFilename(event.target.value)} aria-label={t('اسم ملف CSV', 'CSV filename')} /><span>.csv</span></div></label>
          <fieldset className="encoding-picker">
            <legend>{t('ترميز الملف', 'File encoding')}</legend>
            <label><input type="radio" name="csv-encoding" value="windows-1256" checked={csvEncoding === 'windows-1256'} onChange={() => setCsvEncoding('windows-1256')} />{t('Windows-1256 · دون BOM (قديم)', 'Windows-1256 · no BOM (legacy)')}</label>
            <label><input type="radio" name="csv-encoding" value="utf-8-bom" checked={csvEncoding === 'utf-8-bom'} onChange={() => setCsvEncoding('utf-8-bom')} />{t('UTF-8 · مع BOM (حديث)', 'UTF-8 · with BOM (modern)')}</label>
          </fieldset>
          <button className="gold-button export-button" onClick={requestExport}><ArrowDownToLine size={17} />{t('إنشاء وتنزيل CSV', 'Generate & download CSV')}</button>
          {exportError && <p className="report-error export-error" role="alert"><AlertTriangle size={14} />{exportError}</p>}
        </section>
      </div>}
      {workbook && step === 3 && <motion.section className="glass success-panel success-pop" initial={{ opacity: 0, scale: 0.82 }} animate={{ opacity: 1, scale: 1 }} transition={{ type: 'spring', stiffness: 210, damping: 20 }}><span className="success-icon"><CheckCircle2 size={33} /></span><span className="eyebrow">{exportKind === 'draft' ? t('مسودة مراجعة', 'REVIEW DRAFT') : t('تم التحويل بنجاح', 'CONVERSION COMPLETE')}</span>
        <h3>{exportKind === 'draft' ? t('هذه مسودة مراجعة وليست جاهزة للرفع.', 'This is a review draft, not ready to upload.') : t('ملفك جاهز للعمل.', 'Your file is ready to use.')}</h3>
        <p>{exportKind === 'draft'
          ? t('بقيت خانات مطلوبة أو أخطاء. أكمل الإجابات ثم أنشئ الملف النهائي؛ لا ترفع هذه المسودة إلى بوابة الضرائب.', 'Required answers or errors remain. Complete them and create the final file; do not upload this draft to the tax portal.')
          : t('تم تنزيل نسخة CSV على جهازك. ملف Excel الأصلي لم يغادر متصفحك.', 'A CSV copy has been downloaded. Your original workbook never left your browser.')}</p>
        <div className="success-filename"><FileSpreadsheet size={17} /><b>{outputFiles[0]?.filename ?? safeFilename(filename, workbook.direction)}</b><span>{workbook.rows.length} {t('صف', 'rows')}</span></div>
        {outputFiles.length > 1 && <div className="csv-download-list">
          <p>{t(`الملف مقسّم إلى ${outputFiles.length} أجزاء، وحجم كل جزء لا يتجاوز 3 ميجابايت:`, `The export is split into ${outputFiles.length} parts, each at most 3 MB:`)}</p>
          {outputFiles.map((file, index) => <div key={file.filename}>
            <span>{file.filename}</span>
            <button type="button" className="text-button" onClick={() => downloadCsvPart(file)}><ArrowDownToLine size={14} />{t('تنزيل', 'Download')} {index + 1}</button>
          </div>)}
        </div>}
        <div className="success-actions">
          <button className="gold-button" onClick={() => { if (outputFiles[0]) downloadCsvPart(outputFiles[0]); }}><ArrowDownToLine size={16} />{t('تنزيل مرة أخرى', 'Download again')}</button>
          {exportKind === 'draft'
            ? <button className="text-button" onClick={() => setStep(2)}>{t('العودة لإكمال البيانات', 'Return to complete the answers')} <ArrowLeft size={15} /></button>
            : <button className="text-button" onClick={startOver}>{t('تحويل ملف آخر', 'Convert another file')} <ArrowLeft size={15} /></button>}
        </div>
      </motion.section>}
      {confirm && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) { setConfirm(false); setConfirmMode('warnings'); } }}>
        <section className="confirm-modal glass" role="alertdialog" aria-modal="true" aria-labelledby="confirm-title" aria-describedby="confirm-description"><span className="modal-icon"><AlertTriangle size={22} /></span><button className="modal-close" onClick={() => { setConfirm(false); setConfirmMode('warnings'); }} aria-label={t('إغلاق', 'Close')}><X size={18} /></button>
          <h3 id="confirm-title">{confirmMode === 'draft' ? t('الملف غير مكتمل', 'The file is incomplete') : t('راجع التنبيهات قبل التصدير', 'Review warnings before exporting')}</h3>
          <p id="confirm-description">
            {confirmMode === 'draft'
              ? t(`توجد خانات مطلوبة فارغة أو أخطاء. يمكنك تنزيل مسودة باسم ${workbook?.direction ?? 'sales'}_review_incomplete.csv للمراجعة فقط، لكنها لن تُعرض كملف جاهز للرفع.`, `Required fields are still blank or errors remain. You can download ${workbook?.direction ?? 'sales'}_review_incomplete.csv as a review draft only; it will not be presented as ready to upload.`)
              : t(`يوجد ${workbook?.issues.filter((issue) => issue.severity === 'warning').length ?? 0} تنبيه. القيم المالية الموجودة ستبقى كما هي ولن نعيد حسابها. راجع التنبيهات قبل اعتماد الملف النهائي.`, `There are ${workbook?.issues.filter((issue) => issue.severity === 'warning').length ?? 0} warnings. Existing financial values will be preserved, not recalculated. Review them before approving the final file.`)}
          </p>
          <div className="modal-actions">
            <button className="text-button" onClick={() => { setConfirm(false); setConfirmMode('warnings'); }}>{t('إلغاء والعودة للمراجعة', 'Cancel and return to review')}</button>
            <button className="gold-button" onClick={() => exportCsv(confirmMode === 'draft' ? 'draft' : 'final')}>
              {confirmMode === 'draft' ? t('تنزيل مسودة غير جاهزة للرفع', 'Download incomplete review draft') : t('اعتماد القيم الحالية وإنشاء الملف النهائي', 'Approve current values and create final CSV')}
            </button>
          </div>
        </section>
      </div>}
      {showIosHelp && <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) setShowIosHelp(false); }}>
        <section className="confirm-modal glass" role="dialog" aria-modal="true" aria-labelledby="ios-install-title">
          <button className="modal-close" onClick={() => setShowIosHelp(false)} aria-label={t('إغلاق', 'Close')}><X size={18} /></button>
          <span className="modal-icon"><ArrowDownToLine size={22} /></span>
          <h3 id="ios-install-title">{t('تثبيت التطبيق على iPhone أو iPad', 'Install on iPhone or iPad')}</h3>
          <p>{t('من Safari، اضغط زر المشاركة ثم اختر «إضافة إلى الشاشة الرئيسية». افتح التطبيق لاحقاً من أيقونته ليعمل كتطبيق مستقل.', 'In Safari, tap Share, then choose “Add to Home Screen”. Open it later from its icon to use it like an app.')}</p>
          <div className="modal-actions"><button className="gold-button" onClick={() => setShowIosHelp(false)}>{t('حسناً', 'Got it')}</button></div>
        </section>
      </div>}
    </section>
    <footer className="page-footer">
      <div className="footer-badges">
        <span><WifiOff size={13} />{t('يعمل بدون إنترنت', 'Works offline')}</span>
        <span><LockKeyhole size={13} />{t('بياناتك لا تُرفع', 'Your files stay local')}</span>
        <span><ShieldCheck size={13} />{t('مهيأ لملفات الضرائب المصرية', 'Prepared for Egyptian tax workflows')}</span>
      </div>
      <span><span className="footer-mark">ض</span>{t('محول فواتير الضرائب', 'Tax Invoice Converter')}</span>
      <button onClick={startOver}>{t('بدء من جديد', 'Start over')} <ArrowRight size={13} /></button>
    </footer>
  </main>;
}
export default App;
