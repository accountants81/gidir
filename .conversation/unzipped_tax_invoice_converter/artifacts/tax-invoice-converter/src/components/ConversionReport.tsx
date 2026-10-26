import { useMemo } from 'react';
import { AlertTriangle, FileCheck2 } from 'lucide-react';
import type { CsvEncoding } from '@/lib/local-converter';
import { generateTaxCSVParts } from '@/lib/local-converter';
import { isColumnRequired } from '@/lib/schemas';
import type { ParsedWorkbook } from '@/types/converter';

const CODE_COLUMNS = [0, 1, 2, 13, 14];
const TOTAL_COLUMNS = [
  { column: 19, labelAr: 'إجمالي المبلغ الصافي', labelEn: 'Net total' },
  { column: 20, labelAr: 'إجمالي الضريبة', labelEn: 'Tax total' },
  { column: 21, labelAr: 'الإجمالي النهائي', labelEn: 'Grand total' },
];

export function ConversionReport({
  workbook,
  encoding,
  english,
}: {
  workbook: ParsedWorkbook;
  encoding: CsvEncoding;
  english: boolean;
}) {
  const report = useMemo(() => {
    let bytes = 0;
    let parts = 1;
    let encodingError = '';
    try {
      const csvParts = generateTaxCSVParts(workbook.headers, workbook.rows, encoding);
      bytes = csvParts.reduce((sum, part) => sum + part.byteLength, 0);
      parts = csvParts.length;
    } catch (error) {
      encodingError = error instanceof Error ? error.message : 'CSV encoding failed.';
    }
    const missingRequired = workbook.rows.reduce((count, row) =>
      count + row.reduce((rowCount, _value, column) => rowCount + (isColumnRequired(workbook.direction, row, column) && !row[column]?.trim() ? 1 : 0), 0), 0);
    const codeValues = CODE_COLUMNS.map((column) => ({
      column,
      values: [...new Set(workbook.rows.map((row) => row[column]?.trim()).filter(Boolean))].sort(),
    }));
    const totals = TOTAL_COLUMNS.map(({ column, ...labels }) => ({
      column,
      ...labels,
      value: workbook.rows.reduce((sum, row) => {
        const value = row[column]?.trim() ?? '';
        return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value) ? sum + Number(value) : sum;
      }, 0),
    }));
    const invoiceCount = new Set(workbook.rows.map((row) => row[3]?.trim()).filter(Boolean)).size;
    return { bytes, parts, encodingError, missingRequired, codeValues, totals, invoiceCount };
  }, [workbook, encoding]);

  const money = (value: number) => new Intl.NumberFormat(english ? 'en-US' : 'ar-EG', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  }).format(value);
  const encodingLabel = encoding === 'windows-1256'
    ? (english ? 'Windows-1256 · no BOM' : 'Windows-1256 · دون BOM')
    : (english ? 'UTF-8 · with BOM' : 'UTF-8 · مع BOM');
  const byteLabel = report.bytes < 1024
    ? `${report.bytes} ${english ? 'bytes' : 'بايت'}`
    : `${(report.bytes / 1024).toFixed(2)} KB`;

  return <section className="conversion-report glass" aria-label={english ? 'Pre-export report' : 'تقرير ما قبل التحويل'}>
    <div className="conversion-report-heading">
      <FileCheck2 size={18} />
      <div><h3>{english ? 'Pre-export report' : 'تقرير ما قبل التحويل'}</h3>
        <p>{english ? 'Review the output format and values before downloading.' : 'راجع تنسيق الملف والقيم قبل التنزيل.'}</p></div>
    </div>
    <div className="report-metrics">
      <div><span>{english ? 'Data rows' : 'صفوف البيانات'}</span><b>{workbook.rows.length}</b></div>
      <div><span>{english ? 'Invoices' : 'عدد الفواتير'}</span><b>{report.invoiceCount}</b></div>
      <div><span>{english ? 'Columns' : 'عدد الأعمدة'}</span><b>{workbook.headers.length} / 22</b></div>
      <div><span>{english ? 'Encoding' : 'الترميز'}</span><b>{encodingLabel}</b></div>
      <div><span>{english ? 'Estimated size' : 'الحجم المتوقع'}</span><b>{report.encodingError ? '—' : byteLabel}</b></div>
      <div><span>{english ? 'CSV files' : 'ملفات CSV'}</span><b>{report.parts}</b></div>
      <div><span>{english ? 'Missing required cells' : 'الخانات المطلوبة الفارغة'}</span><b className={report.missingRequired ? 'report-warning' : ''}>{report.missingRequired}</b></div>
      <div><span>{english ? 'Remaining findings' : 'المشكلات والتنبيهات المتبقية'}</span><b className={workbook.issues.length ? 'report-warning' : ''}>{workbook.issues.length}</b></div>
    </div>
    {report.encodingError && <p className="report-error" role="alert"><AlertTriangle size={14} />{report.encodingError}</p>}
    <div className="report-details">
      <div>
        <h4>{english ? 'Codes present in the file' : 'الأكواد الموجودة بالملف'}</h4>
        {report.codeValues.map(({ column, values }) => <p key={column}>
          <b>{english ? ['Document type', 'Tax type', 'Table-goods type', 'Statement type', 'Item type'][CODE_COLUMNS.indexOf(column)] : ['نوع المستند', 'نوع الضريبة', 'نوع سلع الجدول', 'نوع البيان', 'نوع السلعة'][CODE_COLUMNS.indexOf(column)]}:</b>
          {' '}{values.length ? values.join(', ') : '—'}
        </p>)}
      </div>
      <div>
        <h4>{english ? 'Financial totals as entered' : 'إجماليات القيم المالية الموجودة'}</h4>
        {report.totals.map((total) => <p key={total.column}>
          <b>{english ? total.labelEn : total.labelAr}:</b> {money(total.value)}
        </p>)}
      </div>
    </div>
    {workbook.normalizations.map((item) => <p className="report-normalization" key={item.code}>
      <AlertTriangle size={14} />{english ? item.reason : item.reasonAr}
    </p>)}
  </section>;
}
