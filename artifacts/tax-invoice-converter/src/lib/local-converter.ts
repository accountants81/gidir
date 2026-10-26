import type * as XLSXTypes from 'xlsx';
import type { Direction, FormulaCell, ParsedWorkbook, ValidationIssue, NormalizationNotice } from '@/types/converter';
import { ENUM_INDEXES, FINANCIAL_INDEXES, HEADERS, ID_INDEXES, requiredIndexesForRow, enumValues } from './schemas';

const plainNumber = (n: number) => {
  if (!Number.isFinite(n)) return '';
  const text = n.toString();
  if (!/[eE]/.test(text)) return text;
  return expandScientificText(text) ?? '';
};

function expandScientificText(value: string): string | null {
  const match = value.trim().match(/^([+-]?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/);
  if (!match) return null;
  const sign = match[1] === '-' ? '-' : '';
  const whole = match[2];
  const fraction = match[3] ?? '';
  const digits = `${whole}${fraction}`.replace(/^0+(?=\d)/, '');
  const point = whole.length + Number(match[4]);
  if (!Number.isFinite(point)) return null;
  const expanded = point <= 0
    ? `0.${'0'.repeat(-point)}${digits}`
    : point >= digits.length
      ? `${digits}${'0'.repeat(point - digits.length)}`
      : `${digits.slice(0, point)}.${digits.slice(point)}`;
  return `${sign}${expanded}`;
}

const plainNumericPattern = /^[+-]?(?:(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
export const isValidNumericText = (value: string) => plainNumericPattern.test(value.trim());

const numericValue = (value: string): number | null => {
  if (!isValidNumericText(value)) return null;
  const parsed = Number(value.replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
};

const formatDateParts = (year: number, month: number, day: number): string | null => {
  if (year < 1900 || year > 2100 || month < 1 || month > 12) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(year).padStart(4, '0')}`;
};

export function normalizeNumber(value: unknown): string {
  const source = String(value ?? '').trim();
  const expanded = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)[eE][+-]?\d+$/.test(source)
    ? expandScientificText(source)
    : null;
  if (expanded !== null) return expanded;
  if (!isValidNumericText(source)) return '';
  const parsed = Number(source.replace(/,/g, ''));
  return Number.isFinite(parsed) ? plainNumber(parsed) : '';
}

export function normalizeId(value: unknown): string {
  const text = String(value ?? '').trim();
  const expanded = expandScientificText(text);
  if (expanded !== null) return expanded;
  return text.replace(/\.0$/, '');
}

function cellText(cell: XLSXTypes.CellObject | undefined): string {
  if (!cell || cell.v === null || cell.v === undefined) return '';
  if (cell.t === 'n') {
    const number = Number(cell.v);
    if (cell.z && String(cell.z).includes('%') && cell.w) return String(cell.w);
    if (cell.z && /^0+$/.test(String(cell.z)) && cell.w) return String(cell.w);
    return plainNumber(number);
  }
  if (cell.t === 'd') {
    const d = new Date(cell.v as string | number);
    return Number.isNaN(d.getTime())
      ? String(cell.v)
      : `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
  }
  return String(cell.v);
}

const nonEmpty = (row: string[]) => row.some((cell) => cell.trim() !== '');
const normalizeHeaderForMatch = (header: string) => header.trim().replace(/\s+/g, ' ').replace(/\s*([/()])\s*/g, '$1');
const enumFieldLabels: Record<number, { en: string; ar: string }> = {
  0: { en: 'Document type', ar: 'نوع المستند' },
  1: { en: 'Tax type', ar: 'نوع الضريبة' },
  2: { en: 'Table-goods type', ar: 'نوع سلع الجدول' },
  13: { en: 'Statement type', ar: 'نوع البيان' },
  14: { en: 'Item type', ar: 'نوع السلعة' },
};

function excelColumnLabel(zeroBasedIndex: number): string {
  let value = zeroBasedIndex + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function matchesTemplate(headers: string[], direction: Direction): boolean {
  const expectedHeaders = HEADERS[direction];
  if (headers.length !== expectedHeaders.length) return false;
  return expectedHeaders.every((header, index) => {
    const actual = normalizeHeaderForMatch(headers[index] ?? '');
    const expected = normalizeHeaderForMatch(header);
    if (actual === expected) return true;
    if (direction === 'sales' && index === 14) return actual.startsWith(expected.replace(/\)$/, ''));
    return false;
  });
}

export async function getWorkbookSheetNames(file: File): Promise<string[]> {
  const XLSX = await import('xlsx');
  if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('Please choose an Excel workbook (.xlsx or .xls).');
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true, raw: true });
  if (!workbook.SheetNames.length) throw new Error('This workbook does not contain a worksheet.');
  return workbook.SheetNames;
}

interface ValidationResult {
  rows: string[][];
  issues: ValidationIssue[];
  normalizations: NormalizationNotice[];
}

function validTaxId(value: string): boolean {
  return /^\d{9}$/.test(value) && !/^0{9}$/.test(value);
}

function validateRows(
  headers: string[],
  inputRows: string[][],
  direction: Direction,
  sourceRowNumbers: number[],
  applyDefaults: boolean,
): ValidationResult {
  const rows = inputRows.map((row) => [...row]);
  const issues: ValidationIssue[] = [];
  let tableGoodsDefaults = 0;
  let financialNormalizations = 0;
  const push = (
    rowNumber: number,
    column: number,
    severity: ValidationIssue['severity'],
    reason: string,
    reasonAr?: string,
  ) => {
    issues.push({
      rowNumber,
      header: headers[column] ?? `Column ${column + 1}`,
      columnIndex: column,
      cellReference: `${excelColumnLabel(column)}${rowNumber}`,
      severity,
      reason,
      ...(reasonAr ? { reasonAr } : {}),
    });
  };

  rows.forEach((row, index) => {
    const rowNumber = sourceRowNumbers[index] ?? index + 2;
    ID_INDEXES.forEach((column) => {
      if (row[column]) row[column] = normalizeId(row[column]);
    });

    if (applyDefaults && row[1]?.trim() === '1' && !row[2]?.trim()) {
      row[2] = '0';
      tableGoodsDefaults += 1;
    }

    requiredIndexesForRow(direction, row).forEach((column) => {
      if (!row[column]?.trim()) {
        push(rowNumber, column, 'error', 'Required value is missing.', 'هذه الخانة مطلوبة.');
      }
    });

    if (row[1]?.trim() === '2' && !row[2]?.trim()) {
      push(rowNumber, 2, 'error', 'Table-goods type is required when tax type is 2; choose table one (1) or table two (2).', 'نوع سلع الجدول مطلوب عند اختيار نوع الضريبة 2؛ حدّد جدول أول (1) أو جدول ثانٍ (2).');
    }
    if (row[1]?.trim() === '2' && row[2]?.trim() === '0') {
      push(rowNumber, 2, 'error', 'Choose table one (1) or table two (2) when the tax type is 2.', 'عند اختيار نوع الضريبة 2، حدّد جدول أول (1) أو جدول ثانٍ (2)، وليس 0.');
    }

    const invoiceNumber = row[3]?.trim() ?? '';
    if (invoiceNumber && !/^\d+$/.test(invoiceNumber)) {
      push(rowNumber, 3, 'error', 'Invoice number must contain English digits only; remove letters, spaces, and hyphens after reviewing the suggested correction.', 'رقم الفاتورة يجب أن يحتوي على أرقام إنجليزية فقط؛ راجع اقتراح إزالة الحروف والمسافات والشرطات قبل التطبيق.');
    } else if (invoiceNumber && /^0+$/.test(invoiceNumber)) {
      push(rowNumber, 3, 'error', 'Invoice number cannot be zero.', 'رقم الفاتورة لا يمكن أن يكون صفراً.');
    }

    const productName = row[11] ?? '';
    const partyName = row[4] ?? '';
    for (const [column, value, label] of [[11, productName, 'Product name'], [4, partyName, direction === 'sales' ? 'Customer name' : 'Supplier name']] as const) {
      if (/\p{Cc}/u.test(value)) {
        push(rowNumber, column, 'error', `${label} contains a control character.`, `${column === 11 ? 'اسم المنتج' : direction === 'sales' ? 'اسم العميل' : 'اسم المورد'} يحتوي على رمز تحكم غير ظاهر.`);
      }
      if (value.includes(',')) {
        push(rowNumber, column, 'warning', `${label} contains a comma. Review it; replace it with a space if the comma is not part of the real name.`, `${column === 11 ? 'اسم المنتج' : direction === 'sales' ? 'اسم العميل' : 'اسم المورد'} يحتوي على فاصلة؛ راجعها واستبدلها بمسافة إذا لم تكن جزءاً حقيقياً من الاسم.`);
      }
    }

    const numericIdColumns = [5, 6, 8, 9] as const;
    numericIdColumns.forEach((column) => {
      const value = row[column]?.trim() ?? '';
      if (value && !/^\d+$/.test(value)) {
        push(rowNumber, column, 'error', 'Identifier fields must contain English digits only, without spaces or hyphens.', 'حقول الأرقام التعريفية يجب أن تحتوي على أرقام إنجليزية فقط دون مسافات أو شرطات.');
      }
    });

    const taxId = row[5]?.trim() ?? '';
    const taxIdRequired = direction === 'purchases' && row[0]?.trim() !== '4';
    if (taxId && !validTaxId(taxId)) {
      const severity = direction === 'purchases' ? 'error' : 'warning';
      push(rowNumber, 5, severity, 'Tax registration number should be exactly 9 digits and cannot be all zeros; verify the customer value before final submission.', 'رقم التسجيل الضريبي للعميل يُفضّل أن يكون 9 أرقام بالضبط ولا يكون أصفاراً فقط؛ راجع القيمة قبل الإرسال النهائي.');
    } else if (taxIdRequired && !taxId) {
      push(rowNumber, 5, 'error', 'Supplier tax registration number is required for this purchase document.', 'رقم التسجيل الضريبي للمورد مطلوب في مستند المشتريات هذا.');
    }

    const taxRate = row[17] ?? '';
    if (taxRate.trim() && !/^(?:100(?:\.0+)?|\d{1,2}(?:\.\d+)?)$/.test(taxRate.trim())) {
      push(rowNumber, 17, 'error', `Tax rate must be a number from 0 to 100, without spaces or a % sign. Current value: "${taxRate}".`, `فئة الضريبة يجب أن تكون رقماً من 0 إلى 100 دون مسافات أو علامة %. القيمة الحالية: (${taxRate}).`);
    } else if (taxRate.trim() && !['5', '14'].includes(taxRate.trim())) {
      push(rowNumber, 17, 'warning', 'Tax rate is not one of the standard template rates (5 or 14); verify it against the invoice.', 'فئة الضريبة ليست من القيم المعتادة في النموذج (5 أو 14)؛ راجعها مع الفاتورة.');
    }

    FINANCIAL_INDEXES.forEach((column) => {
      const value = row[column]?.trim() ?? '';
      if (!value && !requiredIndexesForRow(direction, row).includes(column)) {
        push(rowNumber, column, 'warning', 'Financial value is empty.', 'الحقل المالي فارغ.');
      } else if (!value) {
        return;
      } else if (!isValidNumericText(value)) {
        push(rowNumber, column, 'error', 'Financial value must be numeric; remove currency, spaces, and incorrectly placed thousands separators.', 'يجب أن تكون القيمة رقمية؛ أزل العملة والمسافات وفواصل الآلاف الموضوعة بشكل غير صحيح.');
      } else {
        const normalized = normalizeNumber(value);
        if (normalized && normalized !== value) financialNormalizations += 1;
        row[column] = normalized;
      }
    });

    const date = row[10] ?? '';
    if (date.trim()) {
      const dateParts = date.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!dateParts) {
        push(rowNumber, 10, 'error', `Date must use DD/MM/YYYY format. Current value: "${date}".`, `صيغة التاريخ يجب أن تكون DD/MM/YYYY. القيمة الحالية: (${date}).`);
      } else if (!formatDateParts(Number(dateParts[3]), Number(dateParts[2]), Number(dateParts[1]))) {
        push(rowNumber, 10, 'error', `Date is not a valid calendar date in DD/MM/YYYY format. Current value: "${date}".`, `التاريخ غير صحيح. استخدم DD/MM/YYYY. القيمة الحالية: (${date}).`);
      }
    }

    ENUM_INDEXES.forEach((column) => {
      const value = row[column] ?? '';
      const allowedValues = enumValues[direction][column] ?? [];
      if (value.trim() && !allowedValues.includes(value.trim())) {
        const field = enumFieldLabels[column] ?? { en: 'Field', ar: 'الحقل' };
        push(rowNumber, column, 'error', `${field.en} must be one of ${allowedValues.join(', ')}. Current value: "${value}".`, `القيمة في خانة ${field.ar} يجب أن تكون إحدى القيم (${allowedValues.join('، ')}). القيمة الحالية: (${value}).`);
      }
    });

    const unitPrice = numericValue(row[16]?.trim() ?? '');
    const quantity = numericValue(row[18]?.trim() ?? '');
    const net = numericValue(row[19]?.trim() ?? '');
    const rate = numericValue(row[17]?.trim() ?? '');
    const tax = numericValue(row[20]?.trim() ?? '');
    const total = numericValue(row[21]?.trim() ?? '');
    const addCalculationWarning = (column: number, calculation: string, reasonAr: string) => {
      push(rowNumber, column, 'warning', `${calculation} differs from the entered value by more than 0.01. Existing financial values were kept unchanged.`, `${reasonAr} الفرق أكبر من 0.01؛ تم الاحتفاظ بالقيم الموجودة دون تعديل.`);
    };
    if (unitPrice !== null && quantity !== null && net !== null && Math.abs(unitPrice * quantity - net) > 0.01) addCalculationWarning(19, 'Net amount does not equal unit price × quantity', 'المبلغ الصافي لا يساوي سعر الوحدة × الكمية');
    if (net !== null && rate !== null && tax !== null && Math.abs((net * rate) / 100 - tax) > 0.01) addCalculationWarning(20, 'Tax amount does not match net amount × tax rate', 'قيمة الضريبة لا تطابق المبلغ الصافي × فئة الضريبة');
    if (net !== null && tax !== null && total !== null && Math.abs(net + tax - total) > 0.01) addCalculationWarning(21, 'Total does not equal net amount + tax', 'الإجمالي لا يساوي المبلغ الصافي + قيمة الضريبة');
  });

  if (!rows.length) {
    push(2, 0, 'error', 'Workbook contains no data rows.', 'الملف لا يحتوي على صفوف فواتير. لا يمكن إنشاء ملف CSV للرفع.');
  } else {
    const invoiceRows = new Map<string, number[]>();
    rows.forEach((row, index) => {
      const invoice = row[3]?.trim();
      if (invoice) invoiceRows.set(invoice, [...(invoiceRows.get(invoice) ?? []), index]);
    });
    invoiceRows.forEach((indexes, invoice) => {
      if (indexes.length < 2) return;
      indexes.forEach((index) => push(sourceRowNumbers[index] ?? index + 2, 3, 'error', `Invoice number "${invoice}" is duplicated in ${indexes.length} rows.`, `رقم الفاتورة (${invoice}) مكرر في ${indexes.length} صفوف.`));
    });
  }

  const normalizations: NormalizationNotice[] = [];
  if (tableGoodsDefaults > 0) normalizations.push({ code: 'general-tax-table-goods-zero', count: tableGoodsDefaults, reason: 'Filled table-goods type with 0 because tax type is 1 (general goods).', reasonAr: `تم وضع 0 في نوع سلع الجدول لـ ${tableGoodsDefaults} صفوف لأن نوع الضريبة هو 1 (سلع عامة).` });
  if (financialNormalizations > 0) normalizations.push({ code: 'financial-format-normalization', count: financialNormalizations, reason: 'Removed valid thousands separators from financial values without changing their numeric value.', reasonAr: `تم توحيد ${financialNormalizations} قيمة مالية بإزالة فواصل آلاف صحيحة دون تغيير قيمتها الرقمية.` });
  return { rows, issues, normalizations };
}

export function revalidateWorkbook(workbook: ParsedWorkbook, inputRows: string[][]): ParsedWorkbook {
  const result = validateRows(workbook.headers, inputRows, workbook.direction, workbook.sourceRowNumbers ?? [], true);
  const carried = workbook.normalizations.filter((item) => !['general-tax-table-goods-zero', 'financial-format-normalization'].includes(item.code));
  return { ...workbook, ...result, normalizations: [...carried, ...result.normalizations] };
}

export async function parseWorkbook(file: File, direction: Direction, selectedSheetName?: string): Promise<ParsedWorkbook> {
  const XLSX = await import('xlsx');
  if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('Please choose an Excel workbook (.xlsx or .xls).');
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true, raw: true });
  const sheetName = selectedSheetName ?? workbook.SheetNames[0];
  if (!sheetName || !workbook.Sheets[sheetName]) throw new Error('This workbook does not contain a worksheet.');
  const sheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  const headers = Array.from({ length: range.e.c - range.s.c + 1 }, (_, offset) => {
    const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + offset })];
    return cell ? String(cell.v ?? '') : '';
  });
  while (headers.length > 0 && !headers[headers.length - 1]?.trim()) headers.pop();
  if (!headers.length || headers.every((header) => !header.trim())) throw new Error('The first row must contain the original column headers.');
  const directionsToCheck: Direction[] = [direction, direction === 'sales' ? 'purchases' : 'sales'];
  const matchedDirection = directionsToCheck.find((candidate) => matchesTemplate(headers, candidate));
  if (!matchedDirection) throw new Error('Workbook columns do not match a supported tax template.');
  const rows: string[][] = [];
  const sourceRowNumbers: number[] = [];
  const formulaCells: FormulaCell[] = [];
  for (let r = range.s.r + 1; r <= range.e.r; r += 1) {
    const row = Array.from({ length: headers.length }, (_, offset) => {
      const cell = sheet[XLSX.utils.encode_cell({ r, c: range.s.c + offset })];
      if (cell?.f) formulaCells.push({ cellReference: `${excelColumnLabel(offset)}${r + 1}`, rowNumber: r + 1, header: headers[offset] ?? `Column ${offset + 1}` });
      return cellText(cell);
    });
    if (nonEmpty(row)) {
      rows.push(row);
      sourceRowNumbers.push(r + 1);
    }
  }
  const result = validateRows(headers, rows, matchedDirection, sourceRowNumbers, true);
  const normalizations = [...result.normalizations];
  if (formulaCells.length) normalizations.push({ code: 'formula-values-used', count: formulaCells.length, reason: 'The workbook contained formulas; cached displayed values were used and formulas were not written to CSV.', reasonAr: `يحتوي الملف على ${formulaCells.length} خلية بمعادلات؛ تم استخدام القيم المحسوبة المخزنة ولن تُكتب المعادلات في CSV.` });
  return { direction: matchedDirection, headers, ...result, normalizations, formulaCells, sourceFilename: file.name, sourceRowNumbers, sourceSheetName: sheetName };
}

function normalizeDateSuggestion(value: string): string | null {
  const parts = value.trim().split(/[.-]/);
  if (parts.length !== 3) return null;
  const [first, second, third] = parts;
  if (first.length === 4) return formatDateParts(Number(first), Number(second), Number(third));
  if (third.length === 2 && Number(second) > 12 && Number(first) <= 12) return formatDateParts(2000 + Number(third), Number(first), Number(second));
  if (third.length === 4 && Number(first) <= 31 && Number(second) <= 12) return formatDateParts(Number(third), Number(second), Number(first));
  return null;
}

export function suggestCorrection(issue: ValidationIssue, value: string): string {
  const trimmed = value.trim();
  if (issue.columnIndex === 3 && /\D/.test(trimmed)) return trimmed.replace(/\D/g, '');
  if ([4, 11].includes(issue.columnIndex)) return value.replace(/[\u0000-\u001f\u007f]/g, '').replace(/,/g, ' ').trim();
  if (issue.columnIndex === 5 && /^[\d\s-]+$/.test(trimmed)) return trimmed.replace(/[\s-]/g, '');
  if (issue.columnIndex === 10) return normalizeDateSuggestion(trimmed) ?? '';
  if (issue.columnIndex === 17 && trimmed.endsWith('%')) return trimmed.slice(0, -1).trim();
  if (FINANCIAL_INDEXES.includes(issue.columnIndex) && isValidNumericText(trimmed)) return normalizeNumber(trimmed);
  return '';
}

export type CsvEncoding = 'windows-1256' | 'utf-8-bom';

const serializeCsvRow = (row: string[]) => row.map((value) => {
  const text = value ?? '';
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}).join(',');

export function generateTaxCSV(headers: string[], rows: string[][]): string {
  if (headers.length !== 22) throw new Error('The CSV must contain exactly 22 columns.');
  const records = [headers, ...rows.filter(nonEmpty)].map((row) => {
    if (row.length !== 22) throw new Error('Every CSV row must contain exactly 22 columns.');
    return serializeCsvRow(row);
  });
  return `${records.join('\r\n')}\r\n`;
}
export const makeCsv = generateTaxCSV;

let windows1256Encoder: Map<string, number> | null = null;
export function encodeWindows1256(value: string): Uint8Array {
  if (!windows1256Encoder) {
    let decoder: TextDecoder;
    try { decoder = new TextDecoder('windows-1256'); } catch { throw new Error('This browser cannot create the legacy Windows-1256 CSV. Choose UTF-8 instead.'); }
    const encoder = new Map<string, number>();
    for (let byte = 0; byte <= 0xff; byte += 1) {
      const decoded = decoder.decode(Uint8Array.of(byte));
      if (decoded.length === 1 && decoded.charCodeAt(0) !== 0xfffd && !encoder.has(decoded)) encoder.set(decoded, byte);
    }
    windows1256Encoder = encoder;
  }
  const bytes: number[] = [];
  for (const character of value) {
    const byte = windows1256Encoder.get(character);
    if (byte === undefined) throw new Error('This CSV contains characters that Windows-1256 cannot represent. Choose UTF-8 or replace the unsupported text.');
    bytes.push(byte);
  }
  return Uint8Array.from(bytes);
}

export function generateTaxCSVBytes(headers: string[], rows: string[][], encoding: CsvEncoding): Uint8Array {
  const csv = generateTaxCSV(headers, rows);
  if (encoding === 'windows-1256') return encodeWindows1256(csv);
  const utf8 = new TextEncoder().encode(csv);
  const withBom = new Uint8Array(utf8.length + 3);
  withBom.set([0xef, 0xbb, 0xbf]);
  withBom.set(utf8, 3);
  return withBom;
}

export const MAX_CSV_PART_BYTES = 3 * 1024 * 1024;
function encodeCsvText(value: string, encoding: CsvEncoding): Uint8Array {
  if (encoding === 'windows-1256') return encodeWindows1256(value);
  const utf8 = new TextEncoder().encode(value);
  const withBom = new Uint8Array(utf8.length + 3);
  withBom.set([0xef, 0xbb, 0xbf]);
  withBom.set(utf8, 3);
  return withBom;
}

export function generateTaxCSVParts(headers: string[], rows: string[][], encoding: CsvEncoding, maxPartBytes = MAX_CSV_PART_BYTES): Uint8Array[] {
  const headerLine = `${serializeCsvRow(headers)}\r\n`;
  const headerBytes = encodeCsvText(headerLine, encoding);
  if (headerBytes.byteLength > maxPartBytes) throw new Error('The CSV header is larger than the 3 MB limit.');
  const parts: Uint8Array[] = [];
  let currentText = headerLine;
  let currentBytes = headerBytes.byteLength;
  let currentRows = 0;
  for (const row of rows.filter(nonEmpty)) {
    const line = `${serializeCsvRow(row)}\r\n`;
    const lineBytes = encoding === 'windows-1256' ? encodeWindows1256(line).byteLength : new TextEncoder().encode(line).byteLength;
    if (currentBytes + lineBytes > maxPartBytes && currentRows > 0) {
      parts.push(encodeCsvText(currentText, encoding));
      currentText = headerLine;
      currentBytes = headerBytes.byteLength;
      currentRows = 0;
    }
    if (currentBytes + lineBytes > maxPartBytes) throw new Error('A single CSV row is larger than the 3 MB limit and cannot be split safely.');
    currentText += line;
    currentBytes += lineBytes;
    currentRows += 1;
  }
  parts.push(encodeCsvText(currentText, encoding));
  return parts;
}

export function safeFilename(value: string, direction: Direction): string {
  const cleaned = value.trim().replace(/\.csv$/i, '').replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^[_-]+|[_-]+$/g, '');
  const now = new Date();
  const date = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  return `${cleaned || `${direction}_final_${date}`}.csv`;
}
