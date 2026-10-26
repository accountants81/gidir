import type { Direction } from '@/types/converter';

export const HEADERS: Record<Direction, string[]> = {
  sales: [
    'نوع المستند (فاتورة 1/ إشعار إضافة 2/إشعار خصم 3/مستهلك نهائي محلي 5/جهة-حكومة 6/مستهلك نهائي أجنبي 7)',
    'نوع الضريبة (سلع عامة 1/سلع جدول 2)',
    'نوع سلع الجدول (لايوجد 0/جدول أولا 1/جدول ثانيا 2)',
    'رقم الفاتورة','اسم العميل','رقم التسجيل الضريبي للعميل','رقم الملف الضريبي للعميل','العنوان','الرقم القومي / رقم جواز السفر','رقم الموبيل','تاريخ الفاتورة','إسم المنتج','كود المنتج',
    'نوع البيان (سلعة 3/خدمة 4/تسويات 5)',
    'نوع السلعة (محلي 1/صادرات 2/آلات ومعدات 5/أجزاء آلات 6/إعفاءات 7/سلع الجدول)',
    'وحدة قياس المنتج','سعر الوحدة','فئة الضريبة (14%/5%)','كمية المنتج','المبلغ الصافي','قيمة الضريبة','إجمالي',
  ],
  purchases: [
    'نوع المستند (فاتورة 1/ إشعار إضافة 2/إشعار خصم 3/إذن إفراج 4)',
    'نوع الضريبة (سلع عامة 1/سلع جدول 2)',
    'نوع سلع الجدول (لايوجد 0/جدول أولا 1/جدول ثانيا 2)',
    'رقم الفاتورة','اسم المورد','رقم التسجيل الضريبي للعميل','رقم الملف الضريبي للعميل','العنوان','الرقم القومي','رقم الموبيل','تاريخ الفاتورة','إسم المنتج','كود المنتج',
    'نوع البيان (محلي 1/مستورد 2/تسويات 5)',
    'نوع السلعة (سلع 3/خدمات 4/آلات ومعدات 5/أجزاء آلات 6/إعفاءات 7)',
    'وحدة قياس المنتج','سعر الوحدة','فئة الضريبة (14%/5%)','كمية المنتج','المبلغ الصافي','قيمة الضريبة','إجمالي',
  ],
};

// These fields are always required in a final upload. Tax registration is added
// conditionally for purchase documents below, while optional identity fields
// remain outside this list.
export const REQUIRED_INDEXES = [0, 1, 2, 3, 4, 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21];
export const FINANCIAL_INDEXES = [16, 18, 19, 20, 21];
export const ID_INDEXES = [3, 5, 6, 8, 9, 12];
export const ENUM_INDEXES = [0, 1, 2, 13, 14];

export function requiredIndexesForRow(direction: Direction, row: string[]): number[] {
  if (direction === 'purchases' && row[0]?.trim() !== '4') return [...REQUIRED_INDEXES, 5];
  return REQUIRED_INDEXES;
}

export function isColumnRequired(direction: Direction, row: string[], column: number): boolean {
  return requiredIndexesForRow(direction, row).includes(column);
}

export const enumValues: Record<Direction, Record<number, string[]>> = {
  sales: {
    0: ['1', '2', '3', '5', '6', '7'],
    1: ['1', '2'],
    2: ['0', '1', '2'],
    13: ['3', '4', '5'],
    14: ['1', '2', '5', '6', '7'],
  },
  purchases: {
    0: ['1', '2', '3', '4'],
    1: ['1', '2'],
    2: ['0', '1', '2'],
    13: ['1', '2', '5'],
    14: ['3', '4', '5', '6', '7'],
  },
};
