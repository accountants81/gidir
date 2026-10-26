import { Check, PencilLine } from 'lucide-react';
import { useMemo, useState } from 'react';
import { suggestCorrection } from '@/lib/local-converter';
import type { ParsedWorkbook, ValidationIssue } from '@/types/converter';

function issueText(issue: ValidationIssue, english: boolean): string {
  return english ? issue.reason : issue.reasonAr ?? issue.reason;
}

export function DataCorrectionsEditor({
  workbook,
  english,
  onApply,
  onRemoveRows,
}: {
  workbook: ParsedWorkbook;
  english: boolean;
  onApply: (issue: ValidationIssue, value: string) => void;
  onRemoveRows: (rowIndexes: number[], reasonAr: string, reason: string) => void;
}) {
  const issues = useMemo(() => {
    const seen = new Set<string>();
    return workbook.issues.filter((issue) => {
      const rowIndex = workbook.sourceRowNumbers?.indexOf(issue.rowNumber) ?? -1;
      const current = rowIndex >= 0 ? workbook.rows[rowIndex]?.[issue.columnIndex] ?? '' : '';
      if (!current.trim()) return false;
      const key = `${issue.cellReference}:${issue.reason}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }, [workbook]);
  const [values, setValues] = useState<Record<string, string>>({});

  if (!issues.length) return null;

  return <section className="glass corrections-panel" aria-label={english ? 'Corrections for invalid existing values' : 'تصحيح القيم الموجودة'}>
    <div className="corrections-heading">
      <div><h3><PencilLine size={17} />{english ? 'Correct existing values' : 'صحّح القيم الموجودة'}</h3>
        <p>{english ? 'Nothing is changed automatically. Review the current value and apply a correction explicitly.' : 'لن نغيّر أي قيمة موجودة تلقائياً. راجع القيمة الحالية ثم طبّق التصحيح صراحةً.'}</p></div>
      <span className="missing-status needs-answers">{issues.length} {english ? 'findings' : 'مخالفة'}</span>
    </div>
    <div className="correction-list">
      {issues.map((issue) => {
        const rowIndex = workbook.sourceRowNumbers?.indexOf(issue.rowNumber) ?? -1;
        const current = rowIndex >= 0 ? workbook.rows[rowIndex]?.[issue.columnIndex] ?? '' : '';
        const key = issue.cellReference;
        const suggested = suggestCorrection(issue, current);
        const value = values[key] ?? (suggested || current);
        const canRemovePurchaseRows = workbook.direction === 'purchases' && issue.columnIndex === 5 && issue.severity === 'error' && rowIndex >= 0;
        const invoiceNumber = rowIndex >= 0 ? workbook.rows[rowIndex]?.[3]?.trim() : '';
        const supplierKey = rowIndex >= 0 ? `${workbook.rows[rowIndex]?.[4]?.trim() ?? ''}|${workbook.rows[rowIndex]?.[7]?.trim() ?? ''}` : '';
        const invoiceRows = canRemovePurchaseRows && invoiceNumber
          ? workbook.rows.map((row, index) => row[3]?.trim() === invoiceNumber ? index : -1).filter((index) => index >= 0)
          : rowIndex >= 0 ? [rowIndex] : [];
        const supplierRows = canRemovePurchaseRows
          ? workbook.rows.map((row, index) => `${row[4]?.trim() ?? ''}|${row[7]?.trim() ?? ''}` === supplierKey ? index : -1).filter((index) => index >= 0)
          : [];
        return <article className={`correction-item ${issue.severity}`} key={`${key}-${issue.reason}`}>
          <div className="correction-meta"><b>{issue.severity === 'error' ? (english ? 'Error' : 'خطأ') : (english ? 'Warning' : 'تنبيه')}</b><span>{english ? 'Cell' : 'الخانة'} <bdi dir="ltr">{issue.cellReference}</bdi> · {issue.header}</span><small>{issueText(issue, english)}</small></div>
          <div className="correction-values"><label><span>{english ? 'Current value' : 'القيمة الحالية'}</span><code dir="auto">{current}</code></label>
            <label><span>{english ? 'Corrected value' : 'القيمة بعد التصحيح'}</span><input value={value} onChange={(event) => setValues((old) => ({ ...old, [key]: event.target.value }))} aria-label={`${english ? 'Corrected value' : 'القيمة بعد التصحيح'} ${issue.cellReference}`} /></label>
            <button type="button" className="text-button apply-missing" disabled={!value.trim()} onClick={() => { onApply(issue, value.trim()); setValues((old) => { const next = { ...old }; delete next[key]; return next; }); }}><Check size={14} />{english ? 'Apply correction' : 'تطبيق التصحيح'}</button>
            {canRemovePurchaseRows && <div className="correction-removal-actions">
              <button type="button" className="text-button danger-button" onClick={() => {
                const message = english ? `Remove ${invoiceRows.length} row(s) for invoice ${invoiceNumber || '(blank)'}?` : `هل تريد حذف ${invoiceRows.length} صف من الفاتورة ${invoiceNumber || '(بدون رقم)'}؟`;
                if (window.confirm(message)) onRemoveRows(invoiceRows, `تم حذف ${invoiceRows.length} صف من الفاتورة ${invoiceNumber || '(بدون رقم)'} بناءً على تأكيد المستخدم.`, `Removed ${invoiceRows.length} row(s) for invoice ${invoiceNumber || '(blank)'} after user confirmation.`);
              }}>{english ? 'Remove this invoice' : 'حذف هذه الفاتورة'}</button>
              {supplierRows.length > invoiceRows.length && <button type="button" className="text-button danger-button" onClick={() => {
                const message = english ? `Remove all ${supplierRows.length} row(s) for this supplier?` : `هل تريد حذف كل فواتير هذا المورد وعددها ${supplierRows.length} صف؟`;
                if (window.confirm(message)) onRemoveRows(supplierRows, `تم حذف كل فواتير المورد وعددها ${supplierRows.length} صف بناءً على تأكيد المستخدم.`, `Removed all ${supplierRows.length} row(s) for this supplier after user confirmation.`);
              }}>{english ? 'Remove all supplier invoices' : 'حذف كل فواتير المورد'}</button>}
            </div>}
          </div>
        </article>;
      })}
    </div>
  </section>;
}
