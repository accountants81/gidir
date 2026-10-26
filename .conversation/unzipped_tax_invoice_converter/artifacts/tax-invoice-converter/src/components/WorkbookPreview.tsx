import { AlertTriangle, CheckCircle2, CircleAlert, FileSpreadsheet } from 'lucide-react';
import type { ParsedWorkbook } from '@/types/converter';

const reasonAr: Record<string, string> = {
  'Required value is missing.': 'قيمة مطلوبة مفقودة.',
  'Financial value is empty.': 'الحقل المالي فارغ.',
  'Financial value must be numeric; remove currency and thousands separators.': 'يجب أن تكون القيمة رقمية؛ أزل رموز العملة وفواصل الآلاف.',
  'Tax ID is not exactly 9 digits.': 'الرقم الضريبي ليس مكوّناً من 9 أرقام بالضبط.',
  'Value is outside the recognized options in this field.': 'القيمة خارج الخيارات المعروفة لهذا الحقل.',
};
export function WorkbookPreview({ workbook, english }: { workbook: ParsedWorkbook; english: boolean }) {
  const errors = workbook.issues.filter((issue) => issue.severity === 'error').length;
  const warnings = workbook.issues.length - errors;
  return <section className="glass preview-card fade-up" aria-label={english ? 'Workbook preview' : 'معاينة المصنف'}>
    <div className="file-summary">
      <span className="file-mark"><FileSpreadsheet size={20} /></span>
      <div className="file-names"><strong dir="auto">{workbook.sourceFilename}</strong><span>{workbook.rows.length} {english ? 'data rows' : 'صف بيانات'} · {workbook.headers.length} {english ? 'columns' : 'عمود'}</span></div>
      <div className="issue-pills">
        {errors > 0 && <span className="issue-pill error"><CircleAlert size={13} />{errors} {english ? 'errors' : 'أخطاء'}</span>}
        {warnings > 0 && <span className="issue-pill warning"><AlertTriangle size={13} />{warnings} {english ? 'warnings' : 'تنبيهات'}</span>}
        {!errors && !warnings && <span className="issue-pill okay"><CheckCircle2 size={13} />{english ? 'Looks good' : 'سليم'}</span>}
      </div>
    </div>
    <div className="table-scroll scrollbar-thin" role="region" tabIndex={0} aria-label={english ? 'Scrollable data preview' : 'جدول البيانات قابل للتمرير'}>
      <table>
        <thead><tr><th className="row-number">#</th>{workbook.headers.map((header, i) => <th key={`${header}-${i}`} dir="auto">{header}</th>)}</tr></thead>
        <tbody>{workbook.rows.slice(0, 20).map((row, ri) => {
          const rowNumber = workbook.sourceRowNumbers?.[ri] ?? ri + 2;
          return <tr key={ri}>
          <td className="row-number">{rowNumber}</td>
          {workbook.headers.map((_, ci) => {
            const issue = workbook.issues.find((item) => item.rowNumber === rowNumber && item.columnIndex === ci);
            return <td key={ci} dir="auto" className={issue ? `cell-${issue.severity}` : ''} title={issue?.reason}>{row[ci] || <span className="empty-cell">—</span>}</td>;
          })}
        </tr>;
        })}</tbody>
      </table>
    </div>
    {workbook.rows.length > 20 && <p className="preview-note">{english ? `+${workbook.rows.length - 20} more rows` : `+${workbook.rows.length - 20} صف إضافي`}</p>}
    {workbook.issues.length > 0 && <details open className="issue-details">
      <summary>{english ? 'Review validation details' : 'تفاصيل التحقق من البيانات'} <span>{workbook.issues.length}</span></summary>
      <ul>{workbook.issues.map((issue, i) => <li key={`${issue.rowNumber}-${issue.header}-${i}`} className={issue.severity}>
        <b>{issue.severity === 'error' ? (english ? 'Error' : 'خطأ') : (english ? 'Warning' : 'تنبيه')}</b>
        <span>{english ? `Row ${issue.rowNumber}` : `الصف ${issue.rowNumber}`} · {english ? 'Cell' : 'الخانة'} <bdi dir="ltr">{issue.cellReference}</bdi> · <span dir="auto">{issue.header}</span></span>
        <span>{english ? issue.reason : (issue.reasonAr ?? reasonAr[issue.reason] ?? issue.reason)}</span>
      </li>)}</ul>
    </details>}
  </section>;
}
