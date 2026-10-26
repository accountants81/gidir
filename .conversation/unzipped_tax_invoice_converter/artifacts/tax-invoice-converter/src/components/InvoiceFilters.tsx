import { Filter, Trash2 } from 'lucide-react';
import { useMemo, useState } from 'react';
import type { Direction, ParsedWorkbook } from '@/types/converter';

type FilterSelection = { month: string; year: string; removeCancelled: boolean };

function invoiceNumbers(workbook: ParsedWorkbook, indexes: number[]): string[] {
  return [...new Set(indexes.map((index) => workbook.rows[index]?.[3]?.trim()).filter(Boolean))];
}

export function InvoiceFilters({
  workbook,
  english,
  onApply,
}: {
  workbook: ParsedWorkbook;
  english: boolean;
  onApply: (selection: FilterSelection, keptIndexes: number[]) => void;
}) {
  const [month, setMonth] = useState('');
  const [year, setYear] = useState('');
  const [removeCancelled, setRemoveCancelled] = useState(false);
  const options = useMemo(() => {
    const values = new Set<string>();
    workbook.rows.forEach((row) => {
      const match = row[10]?.trim().match(/^\d{2}\/(\d{2})\/(\d{4})$/);
      if (match) values.add(`${match[2]}-${match[1]}`);
    });
    return [...values].sort();
  }, [workbook.rows]);
  const selected = Boolean(month || year || removeCancelled);
  const keptIndexes = workbook.rows.flatMap((row, index) => {
    const date = row[10]?.trim().match(/^\d{2}\/(\d{2})\/(\d{4})$/);
    const inMonth = !month || (date && date[1] === month);
    const inYear = !year || (date && date[2] === year);
    const isCancelled = row[0]?.trim() === '3';
    return inMonth && inYear && !(removeCancelled && isCancelled) ? [index] : [];
  });
  const removedIndexes = workbook.rows.map((_, index) => index).filter((index) => !keptIndexes.includes(index));
  const affected = invoiceNumbers(workbook, removedIndexes);
  const monthYears = [...new Set(options.map((item) => item.split('-')[0]))].sort();

  return <details className="glass invoice-filters">
    <summary><Filter size={16} />{english ? 'Optional invoice filtering' : 'تصفية اختيارية للفواتير'}<span>{selected ? (english ? 'Review before applying' : 'راجع قبل التطبيق') : (english ? 'No filter applied' : 'لم تُطبّق تصفية')}</span></summary>
    <div className="filter-body">
      <p>{english ? 'Filtering is optional. The original rows stay untouched until you apply the choices.' : 'التصفية اختيارية. لن تتغير الصفوف الأصلية حتى تضغط تطبيق الاختيارات.'}</p>
      <div className="filter-controls">
        <label><span>{english ? 'Month' : 'الشهر'}</span><select value={month} onChange={(event) => setMonth(event.target.value)}><option value="">{english ? 'All months' : 'كل الشهور'}</option>{Array.from({ length: 12 }, (_, index) => { const value = String(index + 1).padStart(2, '0'); return <option key={value} value={value}>{value}</option>; })}</select></label>
        <label><span>{english ? 'Year' : 'السنة'}</span><select value={year} onChange={(event) => setYear(event.target.value)}><option value="">{english ? 'All years' : 'كل السنوات'}</option>{monthYears.map((value) => <option key={value} value={value}>{value}</option>)}</select></label>
        <label className="filter-check"><input type="checkbox" checked={removeCancelled} onChange={(event) => setRemoveCancelled(event.target.checked)} /><span><Trash2 size={14} />{english ? 'Remove document type 3' : 'حذف إشعارات الخصم / المستند 3'}</span></label>
      </div>
      {selected && <div className="filter-preview"><b>{english ? `${affected.length} invoices / ${removedIndexes.length} rows will be removed` : `سيتم حذف ${affected.length} فاتورة / ${removedIndexes.length} صف`}</b>{affected.length > 0 && <small dir="ltr">{affected.slice(0, 12).join(', ')}{affected.length > 12 ? ' …' : ''}</small>}</div>}
      <div className="filter-actions"><button type="button" className="text-button" onClick={() => { setMonth(''); setYear(''); setRemoveCancelled(false); }}>{english ? 'Clear choices' : 'مسح الاختيارات'}</button><button type="button" className="gold-button" disabled={!selected || !keptIndexes.length || !removedIndexes.length} onClick={() => onApply({ month, year, removeCancelled }, keptIndexes)}>{english ? 'Apply choices' : 'تطبيق الاختيارات'}</button></div>
    </div>
  </details>;
}

export type { FilterSelection };
export type FilterDirection = Direction;
