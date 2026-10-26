import { useState } from 'react';
import { Check, CircleHelp } from 'lucide-react';
import type { Direction, ParsedWorkbook } from '@/types/converter';
import { enumValues, isColumnRequired } from '@/lib/schemas';

type FieldType = 'text' | 'number' | 'enum';
type GroupMode = 'invoice' | 'customer' | 'product' | 'row';
type FieldDefinition = {
  column: number;
  labelAr: string;
  labelEn: string;
  type: FieldType;
  group: GroupMode;
};
type FieldGroup = { id: string; rowIndexes: number[]; context: string };
type EnumOption = { value: string; labelAr: string; labelEn: string };

const fields: FieldDefinition[] = [
  { column: 0, labelAr: 'نوع المستند', labelEn: 'Document type', type: 'enum', group: 'invoice' },
  { column: 1, labelAr: 'نوع الضريبة', labelEn: 'Tax type', type: 'enum', group: 'invoice' },
  { column: 3, labelAr: 'رقم الفاتورة', labelEn: 'Invoice number', type: 'text', group: 'invoice' },
  { column: 4, labelAr: 'اسم العميل / المورد', labelEn: 'Customer / supplier name', type: 'text', group: 'customer' },
  { column: 5, labelAr: 'رقم التسجيل الضريبي', labelEn: 'Tax registration number', type: 'text', group: 'customer' },
  { column: 10, labelAr: 'تاريخ الفاتورة', labelEn: 'Invoice date', type: 'text', group: 'invoice' },
  { column: 11, labelAr: 'اسم المنتج', labelEn: 'Product name', type: 'text', group: 'product' },
  { column: 13, labelAr: 'نوع البيان', labelEn: 'Statement type', type: 'enum', group: 'product' },
  { column: 14, labelAr: 'نوع السلعة', labelEn: 'Item type', type: 'enum', group: 'product' },
  { column: 15, labelAr: 'وحدة القياس', labelEn: 'Unit of measure', type: 'text', group: 'product' },
  { column: 16, labelAr: 'سعر الوحدة', labelEn: 'Unit price', type: 'number', group: 'row' },
  { column: 17, labelAr: 'فئة الضريبة', labelEn: 'Tax rate', type: 'number', group: 'product' },
  { column: 18, labelAr: 'كمية المنتج', labelEn: 'Quantity', type: 'number', group: 'row' },
  { column: 19, labelAr: 'المبلغ الصافي', labelEn: 'Net amount', type: 'number', group: 'row' },
  { column: 20, labelAr: 'قيمة الضريبة', labelEn: 'Tax amount', type: 'number', group: 'row' },
  { column: 21, labelAr: 'الإجمالي', labelEn: 'Total', type: 'number', group: 'row' },
];

const documentLabels: Record<Direction, Record<string, [string, string]>> = {
  sales: {
    '1': ['فاتورة', 'Invoice'],
    '2': ['إشعار إضافة', 'Credit note'],
    '3': ['إشعار خصم', 'Debit note'],
    '5': ['مستهلك نهائي محلي', 'Local final consumer'],
    '6': ['جهة / حكومة', 'Organization / government'],
    '7': ['مستهلك نهائي أجنبي', 'Foreign final consumer'],
  },
  purchases: {
    '1': ['فاتورة', 'Invoice'],
    '2': ['إشعار إضافة', 'Credit note'],
    '3': ['إشعار خصم', 'Debit note'],
    '4': ['إذن إفراج', 'Release permit'],
  },
};
const taxLabels: Record<string, [string, string]> = {
  '1': ['سلع عامة', 'General goods'],
  '2': ['سلع جدول', 'Schedule goods'],
};
const statementLabels: Record<Direction, Record<string, [string, string]>> = {
  sales: { '3': ['سلعة', 'Goods'], '4': ['خدمة', 'Service'], '5': ['تسويات', 'Adjustments'] },
  purchases: { '1': ['محلي', 'Domestic'], '2': ['مستورد', 'Imported'], '5': ['تسويات', 'Adjustments'] },
};
const itemLabels: Record<Direction, Record<string, [string, string]>> = {
  sales: {
    '1': ['محلي', 'Domestic'],
    '2': ['صادرات', 'Exports'],
    '5': ['آلات ومعدات', 'Machinery and equipment'],
    '6': ['أجزاء آلات', 'Machine parts'],
    '7': ['إعفاءات', 'Exemptions'],
  },
  purchases: {
    '3': ['سلع', 'Goods'],
    '4': ['خدمات', 'Services'],
    '5': ['آلات ومعدات', 'Machinery and equipment'],
    '6': ['أجزاء آلات', 'Machine parts'],
    '7': ['إعفاءات', 'Exemptions'],
  },
};

function enumOptions(direction: Direction, column: number): EnumOption[] {
  const labels = column === 0
    ? documentLabels[direction]
    : column === 1
      ? taxLabels
      : column === 2
        ? { '1': ['جدول أول', 'Schedule one'], '2': ['جدول ثانٍ', 'Schedule two'] } as Record<string, [string, string]>
        : column === 13
          ? statementLabels[direction]
          : itemLabels[direction];
  return (enumValues[direction][column] ?? []).map((value) => {
    const [labelAr, labelEn] = labels[value] ?? [value, value];
    return { value, labelAr, labelEn };
  });
}

function groupKey(field: FieldDefinition, row: string[], index: number): string {
  const invoice = row[3]?.trim();
  if (field.group === 'row') return `row:${index}`;
  if (field.group === 'invoice') return invoice ? `invoice:${invoice}` : `row:${index}`;

  if (field.group === 'customer') {
    const taxId = row[5]?.trim();
    if (taxId) return `tax:${taxId}`;
    const nameAndAddress = `${row[4]?.trim()}|${row[7]?.trim()}`;
    if (nameAndAddress !== '|') return `customer:${nameAndAddress}`;
    return invoice ? `invoice:${invoice}` : `row:${index}`;
  }

  const productIdentity = row[12]?.trim() || row[11]?.trim();
  if (productIdentity) {
    return `product:${productIdentity}|${row[13]?.trim()}|${row[14]?.trim()}|${row[1]?.trim()}`;
  }
  return invoice ? `invoice:${invoice}` : `row:${index}`;
}

function contextForRows(workbook: ParsedWorkbook, rowIndexes: number[]): string {
  const invoiceNumbers = [...new Set(rowIndexes.map((index) => workbook.rows[index]?.[3]?.trim()).filter(Boolean))];
  const customers = [...new Set(rowIndexes.map((index) => workbook.rows[index]?.[4]?.trim()).filter(Boolean))];
  const rows = rowIndexes.map((index) => workbook.sourceRowNumbers?.[index] ?? index + 2);
  const rowLabel = rows.length <= 4
    ? rows.join('، ')
    : `${rows.length} ${rows.length === 1 ? 'صف' : 'صفوف'}`;
  return invoiceNumbers.length === 1
    ? `فاتورة ${invoiceNumbers[0]} · الصف ${rowLabel}`
    : customers.length === 1
      ? `${customers[0]} · الصف ${rowLabel}`
      : `الصف ${rowLabel}`;
}

function buildGroups(workbook: ParsedWorkbook, field: FieldDefinition): FieldGroup[] {
  const groups = new Map<string, number[]>();
  workbook.rows.forEach((row, index) => {
    if (row[field.column]?.trim()) return;
    const key = groupKey(field, row, index);
    groups.set(key, [...(groups.get(key) ?? []), index]);
  });

  return [...groups.entries()].map(([id, rowIndexes]) => ({
    id,
    rowIndexes,
    context: contextForRows(workbook, rowIndexes),
  }));
}

export function MissingValuesEditor({
  workbook,
  english,
  onApply,
  onApplySingleItemDefaults,
}: {
  workbook: ParsedWorkbook;
  english: boolean;
  onApply: (column: number, value: string, rowIndexes: number[]) => void;
  onApplySingleItemDefaults: (rowIndexes: number[]) => void;
}) {
  const [values, setValues] = useState<Record<string, string>>({});
  const t = (ar: string, en: string) => english ? en : ar;
  const missingFields = fields
    .filter((field) => workbook.rows.some((row) => isColumnRequired(workbook.direction, row, field.column)))
    .map((field) => ({ ...field, groups: buildGroups(workbook, field) }))
    .filter((field) => field.groups.length > 0);

  const tableGoodsField: FieldDefinition = {
    column: 2,
    labelAr: 'نوع سلع الجدول',
    labelEn: 'Table-goods type',
    type: 'enum',
    group: 'invoice',
  };
  const tableGoodsGroups = buildGroups(workbook, tableGoodsField)
    .map((group) => ({
      ...group,
      rowIndexes: group.rowIndexes.filter((index) => workbook.rows[index]?.[1]?.trim() === '2'),
    }))
    .map((group) => ({ ...group, context: contextForRows(workbook, group.rowIndexes) }))
    .filter((group) => group.rowIndexes.length > 0);
  if (tableGoodsGroups.length) {
    missingFields.splice(2, 0, { ...tableGoodsField, groups: tableGoodsGroups });
  }

  const priceOrQuantityRows = workbook.rows.flatMap((row, index) =>
    !row[16]?.trim() || !row[18]?.trim() ? [index] : [],
  );
  const invoiceNumbers = workbook.rows.map((row) => row[3]?.trim() ?? '');
  const oneRowPerInvoice = workbook.rows.length > 0
    && invoiceNumbers.every(Boolean)
    && new Set(invoiceNumbers).size === invoiceNumbers.length;
  const hasNetForMissingPrices = priceOrQuantityRows.every((index) => workbook.rows[index]?.[19]?.trim());
  const allPriceAndQuantityMissing = priceOrQuantityRows.every((index) =>
    !workbook.rows[index]?.[16]?.trim() && !workbook.rows[index]?.[18]?.trim(),
  );
  const canOfferSingleItemRule = oneRowPerInvoice && hasNetForMissingPrices && allPriceAndQuantityMissing && priceOrQuantityRows.length > 0;

  return <section className="glass missing-values" aria-label={english ? 'Questions about missing values' : 'أسئلة القيم الناقصة'}>
    <div className="missing-values-heading">
      <div>
        <h3>{t('أسئلة قبل تجهيز الملف النهائي', 'Questions before preparing the final file')}</h3>
        <p>{t('أكمل القيم الناقصة بنفسك. لن نخمّن بيانات العملاء أو الأسعار أو الكميات.', 'Answer the missing fields yourself. Customer data, prices, and quantities are never guessed.')}</p>
      </div>
      <span className={`missing-status ${missingFields.length ? 'needs-answers' : 'answers-complete'}`}>
        {missingFields.length
          ? t(`${missingFields.reduce((sum, field) => sum + field.groups.length, 0)} سؤال`, `${missingFields.reduce((sum, field) => sum + field.groups.length, 0)} questions`)
          : t('مكتمل', 'Complete')}
      </span>
    </div>

    {canOfferSingleItemRule && <div className="single-item-question">
      <CircleHelp size={17} />
      <div>
        <b>{t('هل كل فاتورة تمثل بندًا واحدًا؟', 'Does each invoice contain one line item?')}</b>
        <p>{t('إذا أكدت ذلك، سنضع الكمية 1 وسعر الوحدة مساويًا للمبلغ الصافي في الصفوف الناقصة فقط.', 'If you confirm, quantity becomes 1 and unit price matches the net amount, only in rows where those values are missing.')}</p>
      </div>
      <button
        type="button"
        className="text-button"
        data-testid="button-confirm-single-item"
        onClick={() => onApplySingleItemDefaults(priceOrQuantityRows)}
      >
        {t('نعم، طبّق على', 'Yes, apply to')} {priceOrQuantityRows.length} {t('صف', 'rows')}
      </button>
    </div>}

    {missingFields.length ? <div className="missing-fields">
      {missingFields.map((field) => <article className="missing-field" key={field.column}>
        <div className="missing-field-heading">
          <label>{english ? field.labelEn : field.labelAr}</label>
          <small>{t(`${field.groups.reduce((sum, group) => sum + group.rowIndexes.length, 0)} خانة ناقصة`, `${field.groups.reduce((sum, group) => sum + group.rowIndexes.length, 0)} blank cells`)}</small>
        </div>
        <div className="missing-group-list">
          {field.groups.map((group) => {
            const inputId = `missing-${field.column}-${group.id}`;
            const key = `${field.column}:${group.id}`;
            const rowLabels = group.rowIndexes
              .map((index) => workbook.sourceRowNumbers?.[index] ?? index + 2);
            const countLabel = rowLabels.length === 1
              ? t(`صف ${rowLabels[0]}`, `Row ${rowLabels[0]}`)
              : t(`${rowLabels.length} صفوف`, `${rowLabels.length} rows`);
            return <div className="missing-group" key={group.id}>
              <label className="missing-context" htmlFor={inputId}>
                <b>{group.context}</b>
                <small>{countLabel}{group.rowIndexes.length > 1 ? ` · ${t('ستُطبّق الإجابة على هذه الصفوف فقط', 'answer applies to these rows only')}` : ''}</small>
              </label>
              {field.type === 'enum' ? <select
                id={inputId}
                aria-label={`${english ? field.labelEn : field.labelAr} · ${group.context}`}
                value={values[key] ?? ''}
                onChange={(event) => setValues((current) => ({ ...current, [key]: event.target.value }))}
              >
                <option value="">{t('اختر الكود', 'Choose a code')}</option>
                {enumOptions(workbook.direction, field.column).map((option) => <option value={option.value} key={option.value}>
                  {option.value} — {english ? option.labelEn : option.labelAr}
                </option>)}
              </select> : <input
                id={inputId}
                type={field.type === 'number' ? 'number' : 'text'}
                step={field.type === 'number' ? 'any' : undefined}
                inputMode={field.type === 'number' ? 'decimal' : undefined}
                aria-label={`${english ? field.labelEn : field.labelAr} · ${group.context}`}
                value={values[key] ?? ''}
                onChange={(event) => setValues((current) => ({ ...current, [key]: event.target.value }))}
                placeholder={field.column === 10 ? 'dd/mm/yyyy' : t('اكتب الإجابة', 'Enter an answer')}
              />}
              <button
                type="button"
                className="text-button apply-missing"
                disabled={!values[key]?.trim()}
                onClick={() => {
                  const value = values[key]?.trim();
                  if (!value) return;
                  onApply(field.column, value, group.rowIndexes);
                  setValues((current) => ({ ...current, [key]: '' }));
                }}
              >
                <Check size={14} />{t('تطبيق الإجابة', 'Apply answer')}
              </button>
            </div>;
          })}
        </div>
      </article>)}
    </div> : <p className="missing-values-complete">{t('لا توجد أسئلة عن الحقول الإلزامية. راجع التنبيهات قبل إنشاء الملف.', 'No required fields are missing. Review any warnings before creating the file.')}</p>}

    <p className="missing-values-note">
      {t('رقم الملف الضريبي والرقم القومي ورقم الموبايل وكود المنتج اختيارية. إذا كانت إجابات الحقل تختلف بين الصفوف، أجب لكل مجموعة أو صف على حدة.', 'Tax file number, national ID, mobile number, and product code are optional. If a value differs by invoice or line, answer each group separately.')}
    </p>
  </section>;
}
