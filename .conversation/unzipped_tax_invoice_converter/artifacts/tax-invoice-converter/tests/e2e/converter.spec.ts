import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { basename, resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import * as XLSX from 'xlsx';

const salesHeaders = [
  'نوع المستند (فاتورة 1/ إشعار إضافة 2/إشعار خصم 3/مستهلك نهائي محلي 5/جهة-حكومة 6/مستهلك نهائي أجنبي 7)',
  'نوع الضريبة (سلع عامة 1/سلع جدول 2)',
  'نوع سلع الجدول (لايوجد 0/جدول أولا 1/جدول ثانيا 2)',
  'رقم الفاتورة',
  'اسم العميل',
  'رقم التسجيل الضريبي للعميل',
  'رقم الملف الضريبي للعميل',
  'العنوان',
  'الرقم القومي / رقم جواز السفر',
  'رقم الموبيل',
  'تاريخ الفاتورة',
  'إسم المنتج',
  'كود المنتج',
  'نوع البيان (سلعة 3/خدمة 4/تسويات 5)',
  'نوع السلعة (محلي 1/صادرات 2/آلات ومعدات 5/أجزاء آلات 6/إعفاءات 7/سلع الجدول)',
  'وحدة قياس المنتج',
  'سعر الوحدة',
  'فئة الضريبة (14%/5%)',
  'كمية المنتج',
  'المبلغ الصافي',
  'قيمة الضريبة',
  'إجمالي',
];

const purchaseHeaders = [
  'نوع المستند (فاتورة 1/ إشعار إضافة 2/إشعار خصم 3/إذن إفراج 4)',
  'نوع الضريبة (سلع عامة 1/سلع جدول 2)',
  'نوع سلع الجدول (لايوجد 0/جدول أولا 1/جدول ثانيا 2)',
  'رقم الفاتورة',
  'اسم المورد',
  'رقم التسجيل الضريبي للعميل',
  'رقم الملف الضريبي للعميل',
  'العنوان',
  'الرقم القومي',
  'رقم الموبيل',
  'تاريخ الفاتورة',
  'إسم المنتج',
  'كود المنتج',
  'نوع البيان (محلي 1/مستورد 2/تسويات 5)',
  'نوع السلعة (سلع 3/خدمات 4/آلات ومعدات 5/أجزاء آلات 6/إعفاءات 7)',
  'وحدة قياس المنتج',
  'سعر الوحدة',
  'فئة الضريبة (14%/5%)',
  'كمية المنتج',
  'المبلغ الصافي',
  'قيمة الضريبة',
  'إجمالي',
];

function validRow(id: number, date = '15/01/2024'): string[] {
  const row = Array<string>(22).fill('');
  Object.assign(row, {
    0: '1',
    1: '1',
    2: '0',
    3: String(202400000 + id),
    4: 'شركة النور',
    5: '012345678',
    6: '0000456',
    7: 'القاهرة',
    8: '00000000000000',
    9: '01000000000',
    10: date,
    11: 'منتج تجريبي',
    12: '0012',
    13: '3',
    14: '1',
    15: 'قطعة',
    16: id === 1 ? '1,200.50' : '1200.50',
    17: '14',
    18: '2',
    19: '2401',
    20: '336.14',
    21: '2737.14',
  });
  return row;
}

function decodeWindows1256(bytes: Buffer): string {
  return new TextDecoder('windows-1256').decode(bytes);
}

function workbookBuffer(headers: string[], rows: Array<Array<string | number | Date | null>>): Buffer {
  const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Transactions');
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }) as Buffer;
}

function legacyWorkbookBuffer(headers: string[], rows: Array<Array<string | number | Date | null>>): Buffer {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...rows]), 'First sheet');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([purchaseHeaders, [ ...validRow(99) ]]), 'Ignored sheet');
  return XLSX.write(workbook, { bookType: 'biff8', type: 'buffer' }) as Buffer;
}

async function uploadWorkbook(page: Page, direction: 'sales' | 'purchases', name: string, buffer: Buffer) {
  await page.getByTestId(`button-select-${direction}`).click();
  const chooserPromise = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'اختيار ملف', exact: true }).click();
  const chooser = await chooserPromise;
  await chooser.setFiles({
    name,
    mimeType: name.toLowerCase().endsWith('.xls') ? 'application/vnd.ms-excel' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer,
  });
}

async function dropWorkbook(page: Page, direction: 'sales' | 'purchases', name: string, buffer: Buffer) {
  await page.getByTestId(`button-select-${direction}`).click();
  await page.locator('.upload-zone').evaluate((zone, file) => {
    const binary = atob(file.base64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    const workbook = new File([bytes], file.name, { type: file.mimeType });
    const transfer = new DataTransfer();
    transfer.items.add(workbook);
    zone.dispatchEvent(new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: transfer }));
    zone.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer }));
  }, {
    name,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: buffer.toString('base64'),
  });
}

test.beforeEach(async ({ page }) => {
  await page.goto('/');
});

test('requires a sales or purchases choice before a file can be selected', async ({ page }) => {
  await expect(page.locator('.upload-zone')).toHaveCount(0);
  await expect(page.locator('input[type="file"]')).toHaveCount(0);
  await page.getByTestId('button-select-purchases').click();
  await expect(page.locator('.upload-zone')).toBeVisible();
  await expect(page.locator('input[type="file"]')).toHaveAttribute('accept', /xlsx.*xls/);
});

test('keeps sales conversion local and downloads an explicitly marked Windows-1256 review draft', async ({ page }) => {
  const nonLocalRequests: string[] = [];
  page.on('request', (request) => {
    if (!['GET', 'HEAD'].includes(request.method())) nonLocalRequests.push(`${request.method()} ${request.url()}`);
  });
  const sourceHeaders = [...salesHeaders];
  sourceHeaders[14] = 'نوع السلعة (محلي 1/صادرات 2/آلات ومعدات 5/أجزاء آلات 6/إعفاءات 7 /  سلع الجدول  مراجعة الإرشادات ) ';
  const valid = validRow(1, '15/01/2024');
  valid[4] = 'شركة "النور"';
  const scientificId = validRow(2, '2024-01-15');
  scientificId[5] = '1.23E+08';
  const dashDate = validRow(3, '15-01-2024');
  const twoDigitDate = validRow(4, '15/1/24');
  const dateObjectRow: Array<string | Date> = [...validRow(5)];
  dateObjectRow[10] = new Date('2024-01-15T00:00:00.000Z');
  const invalidDate = validRow(6, '31-02-2024');
  scientificId[16] = '1.25E+03';
  const invalid = Array<string>(22).fill('');
  invalid[16] = '';
  invalid[19] = 'قيمة غير رقمية';
  await dropWorkbook(page, 'sales', 'sales-fixture.xlsx', workbookBuffer(sourceHeaders, [
    valid,
    Array<string>(22).fill(''),
    scientificId,
    dashDate,
    twoDigitDate,
    dateObjectRow,
    invalidDate,
    invalid,
  ]));

  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
  await expect(page.locator('.file-names span')).toContainText('7 صف بيانات');
  await expect(page.locator('.table-scroll thead th')).toHaveCount(23);
  await expect(page.locator('.issue-pill.error')).toBeVisible();
  await expect(page.locator('.issue-pill.warning')).toBeVisible();
  await expect(page.locator('.issue-details li').filter({ hasText: 'الصف 9' }).first()).toBeVisible();
  await expect(page.locator('.issue-details li').filter({ hasText: 'K7' })).toHaveCount(0);

  await page.getByRole('textbox', { name: 'اسم ملف CSV' }).fill('tax:october');
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  await expect(page.getByRole('alertdialog')).toContainText('غير مكتمل');
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'تنزيل مسودة غير جاهزة للرفع' }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe('sales_review_incomplete.csv');

  const downloadedPath = await download.path();
  expect(downloadedPath).toBeTruthy();
  const bytes = await readFile(downloadedPath!);
  expect([...bytes.subarray(0, 3)]).not.toEqual([0xef, 0xbb, 0xbf]);
  const csv = decodeWindows1256(bytes);
  expect(csv).toContain('"شركة ""النور"""');
  expect(csv).toContain('15/01/2024');
  expect(csv).toContain('2024-01-15');
  expect(csv).toContain('123000000');
  expect(csv).toContain('012345678');
  expect(csv).toContain('0000456');
  expect(csv).toContain('1200.5');
  expect(csv).toContain('1250');
  expect(csv).toContain('31-02-2024');
  expect(csv).toContain('قيمة غير رقمية');
  expect(csv).toContain(sourceHeaders[14]);
  expect(csv.endsWith('\r\n')).toBe(true);
  expect(csv.replace(/\r\n/g, '').includes('\n')).toBe(false);
  const records = csv.split('\r\n').filter(Boolean);
  expect(records).toHaveLength(8);
  expect(nonLocalRequests).toEqual([]);
  await expect(page.getByRole('heading', { name: 'هذه مسودة مراجعة وليست جاهزة للرفع.' })).toBeVisible();
});

test('reports portal validation issues and only permits an incomplete draft until they are fixed', async ({ page }) => {
  const invalidDocument = validRow(1);
  invalidDocument[0] = 'فاتورة';
  const emptyTaxType = validRow(2);
  emptyTaxType[1] = '';
  const invalidDate = validRow(3, '24-10-2026');
  const percentTaxRate = validRow(4);
  percentTaxRate[17] = '14%';

  await uploadWorkbook(page, 'sales', 'portal-validation.xlsx', workbookBuffer(salesHeaders, [
    invalidDocument,
    emptyTaxType,
    invalidDate,
    percentTaxRate,
  ]));

  const issues = page.locator('.issue-details li');
  await expect(issues.filter({ hasText: 'A2' })).toContainText('نوع المستند');
  await expect(issues.filter({ hasText: 'A2' })).toContainText('فاتورة');
  await expect(issues.filter({ hasText: 'B3' })).toContainText('نوع الضريبة');
  await expect(issues.filter({ hasText: 'K4' })).toContainText('24-10-2026');
  await expect(issues.filter({ hasText: 'R5' })).toContainText('14%');

  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  await expect(page.getByRole('alertdialog')).toContainText('خانات مطلوبة فارغة أو أخطاء');
  const draftPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'تنزيل مسودة غير جاهزة للرفع' }).click();
  const draft = await draftPromise;
  expect(draft.suggestedFilename()).toBe('sales_review_incomplete.csv');
  await expect(page.getByRole('heading', { name: 'هذه مسودة مراجعة وليست جاهزة للرفع.' })).toBeVisible();
  await page.getByRole('button', { name: 'العودة لإكمال البيانات' }).click();
  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
});

test('requires approval even when the only finding is a warning', async ({ page }) => {
  const row = validRow(7);
  row[5] = '123';
  await uploadWorkbook(page, 'sales', 'warning-only.xlsx', workbookBuffer(salesHeaders, [row]));

  await expect(page.locator('.issue-pill.error')).toHaveCount(0);
  await expect(page.locator('.issue-pill.warning')).toBeVisible();
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  await expect(page.getByRole('alertdialog')).toBeVisible();
  await expect(page.getByRole('alertdialog')).toContainText('تنبيه');
  await page.getByRole('button', { name: 'إلغاء والعودة للمراجعة' }).click();
  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
  await expect(page.getByRole('heading', { name: 'ملفك جاهز للعمل.' })).not.toBeVisible();
});

test('asks which worksheet to use in a workbook with multiple sheets', async ({ page }) => {
  await uploadWorkbook(page, 'sales', 'legacy.xls', legacyWorkbookBuffer(salesHeaders, [validRow(42)]));
  await expect(page.getByLabel('الملف يحتوي على أكثر من ورقة. اختر الورقة المطلوبة:')).toBeVisible();
  await page.locator('#workbook-sheet').selectOption('First sheet');
  await page.getByRole('button', { name: 'متابعة بهذه الورقة' }).click();
  await expect(page.locator('.file-names span')).toContainText('1 صف بيانات');
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  const download = await downloadPromise;
  const csv = decodeWindows1256(await readFile((await download.path())!));
  expect(csv).toContain('202400042');
  expect(csv).not.toContain('202400099');
});

test('asks about blank required fields and requires confirmation before using the one-item rule', async ({ page }) => {
  const row = validRow(33);
  row[15] = '';
  row[16] = '';
  row[18] = '';
  await uploadWorkbook(page, 'sales', 'missing-fields.xlsx', workbookBuffer(salesHeaders, [row]));

  await expect(page.getByRole('heading', { name: 'أسئلة قبل تجهيز الملف النهائي' })).toBeVisible();
  await expect(page.getByRole('textbox', { name: /وحدة القياس/ })).toBeVisible();
  await page.getByRole('textbox', { name: /وحدة القياس/ }).fill('قطعة');
  await page.locator('.missing-field').filter({ hasText: 'وحدة القياس' }).getByRole('button', { name: 'تطبيق الإجابة' }).click();
  await expect(page.getByTestId('button-confirm-single-item')).toBeVisible();
  await page.getByTestId('button-confirm-single-item').click();
  await expect(page.locator('.missing-status')).toHaveText('مكتمل');
  await expect(page.locator('.report-metrics > div').filter({ hasText: 'الخانات المطلوبة الفارغة' }).locator('b')).toHaveText('0');

  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).not.toContain('review_incomplete');
  await expect(page.getByRole('heading', { name: 'ملفك جاهز للعمل.' })).toBeVisible();
});

test('reads the purchases schema and exports the original Arabic headers', async ({ page }) => {
  const purchase = validRow(21);
  purchase[13] = '1';
  purchase[14] = '3';
  await uploadWorkbook(page, 'purchases', 'purchases-fixture.xlsx', workbookBuffer(purchaseHeaders, [purchase]));
  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
  await expect(page.locator('.table-scroll thead')).toContainText('اسم المورد');
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/^purchases_final_\d{8}\.csv$/);
  const csv = decodeWindows1256(await readFile((await download.path())!));
  expect(csv.split('\r\n')[0]).toContain('اسم المورد');
  expect(csv.split('\r\n')[0]).toContain('نوع المستند (فاتورة 1/ إشعار إضافة 2/إشعار خصم 3/إذن إفراج 4)');
});

test('offers explicit deletion of a purchase invoice with an invalid supplier tax number', async ({ page }) => {
  const purchase = validRow(22);
  purchase[5] = 'BAD-TAX-ID';
  await uploadWorkbook(page, 'purchases', 'purchase-invalid-tax.xlsx', workbookBuffer(purchaseHeaders, [purchase]));
  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
  await expect(page.locator('.corrections-panel')).toContainText('رقم التسجيل الضريبي');
  page.once('dialog', (dialog) => dialog.accept());
  await page.getByRole('button', { name: 'حذف هذه الفاتورة' }).first().click();
  await expect(page.locator('.file-names span')).toContainText('0 صف بيانات');
});

test('filters purchases by the selected month and actual year', async ({ page }) => {
  const october = validRow(31, '15/10/2026');
  const september = validRow(32, '15/09/2026');
  await uploadWorkbook(page, 'purchases', 'purchase-filter.xlsx', workbookBuffer(purchaseHeaders, [october, september]));
  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
  const filters = page.locator('details.invoice-filters');
  await filters.locator('summary').click();
  await expect(filters.locator('select').nth(1).locator('option')).toContainText(['كل السنوات', '2026']);
  await filters.locator('select').nth(0).selectOption('10');
  await filters.locator('select').nth(1).selectOption('2026');
  await filters.locator('input[type=checkbox]').check();
  await expect(filters.locator('.filter-preview')).toContainText('سيتم حذف 1 فاتورة / 1 صف');
  await filters.getByRole('button', { name: 'تطبيق الاختيارات' }).click();
  await expect(page.locator('.file-names span')).toContainText('1 صف بيانات');
  await expect(page.locator('.table-scroll tbody')).toContainText('202400031');
});

test('accepts a workbook dropped onto the upload area', async ({ page }) => {
  await dropWorkbook(page, 'sales', 'dropped.xlsx', workbookBuffer(salesHeaders, [validRow(9)]));
  await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
  await expect(page.locator('.file-names strong')).toHaveText('dropped.xlsx');
  await expect(page.locator('.issue-pill.okay')).toBeVisible();
});

test('rejects unsupported and damaged files with a visible message', async ({ page }) => {
  await page.getByTestId('button-select-sales').click();
  const input = page.locator('input[type="file"]');
  await input.setInputFiles({ name: 'notes.txt', mimeType: 'text/plain', buffer: Buffer.from('not Excel') });
  await expect(page.getByRole('alert')).toContainText('يرجى اختيار ملف Excel');
  await input.setInputFiles({ name: 'damaged.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from('broken workbook') });
  await expect(page.getByRole('alert')).toContainText('عناوين الأعمدة لا تطابق نموذج المبيعات أو المشتريات');
  await expect(page.getByRole('heading', { name: 'ما نوع المعاملة؟' })).toBeVisible();
});

test('opens the supplied sales and purchases workbook samples without changing their headers', async ({ page }) => {
  const samples = [
    { path: resolve(process.cwd(), '../../attached_assets/Sales_ar__2__1792692500962.xlsx'), direction: 'sales' as const },
    { path: resolve(process.cwd(), '../../attached_assets/Purchases_ar__2__1792692500938.xlsx'), direction: 'purchases' as const },
  ];
  test.skip(samples.some((sample) => !existsSync(sample.path)), 'User-provided sample workbooks are not present in this checkout.');
  for (const sample of samples) {
    await page.goto('/');
    const bytes = await readFile(sample.path);
    await uploadWorkbook(page, sample.direction, basename(sample.path), bytes);
    await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
    await expect(page.locator('.file-names span')).toContainText('0 صف بيانات');
    await expect(page.locator('.table-scroll thead th')).toHaveCount(23);
    if (sample.direction === 'sales') {
      await expect(page.locator('.table-scroll thead')).toContainText('مراجعة الإرشادات');
    } else {
      await expect(page.locator('.table-scroll thead')).toContainText('اسم المورد');
    }
  }
});

test('accepts both new workbooks from either section and preserves the detected template headers', async ({ page }) => {
  const samples = [
    {
      path: resolve(process.cwd(), '../../attached_assets/00_1792956374908.xlsx'),
      direction: 'sales' as const,
      identityHeader: 'اسم العميل',
    },
    {
      path: resolve(process.cwd(), '../../attached_assets/رقم_المستند_مشتريات_20261025_2213_1792956374932.xlsx'),
      direction: 'purchases' as const,
      identityHeader: 'اسم المورد',
    },
  ];
  test.skip(samples.some((sample) => !existsSync(sample.path)), 'The newly supplied workbooks are not present in this checkout.');

  for (const sample of samples) {
    for (const selectedDirection of ['sales', 'purchases'] as const) {
      await page.goto('/');
      const bytes = await readFile(sample.path);
      const sourceWorkbook = XLSX.read(bytes, { type: 'buffer', cellDates: true, raw: true });
      const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
      const sourceHeaders = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, range: 'A1:V1', blankrows: false })[0] as string[];
      const sourceRows = XLSX.utils.sheet_to_json<Array<string | number | Date>>(sourceSheet, {
        header: 1,
        raw: true,
        defval: '',
        blankrows: false,
      });
      const sourceDates = sourceRows.slice(1).map((row) => {
        const value = row[10];
        return value instanceof Date
          ? `${String(value.getUTCDate()).padStart(2, '0')}/${String(value.getUTCMonth() + 1).padStart(2, '0')}/${value.getUTCFullYear()}`
          : String(value ?? '').trim();
      }).filter(Boolean);

      await uploadWorkbook(page, selectedDirection, basename(sample.path), bytes);
      await expect(page.getByRole('heading', { name: 'معاينة وتحقق' })).toBeVisible();
      await expect(page.locator('.table-scroll thead th')).toHaveCount(23);
      await expect(page.locator('.table-scroll thead')).toContainText(sample.identityHeader);

      const downloadPromise = page.waitForEvent('download');
      await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
      const exportResult = await Promise.race([
        downloadPromise.then((download) => ({ download })),
        page.getByRole('alertdialog').waitFor({ state: 'visible' }).then(() => ({ needsConfirmation: true as const })),
      ]);
      const download = 'download' in exportResult
        ? exportResult.download
        : await (async () => {
          const confirmedDownload = page.waitForEvent('download');
          const dialog = await page.getByRole('alertdialog').innerText();
          const buttonName = dialog.includes('غير مكتمل')
            ? 'تنزيل مسودة غير جاهزة للرفع'
            : 'اعتماد القيم الحالية وإنشاء الملف النهائي';
          await page.getByRole('button', { name: buttonName }).click();
          return confirmedDownload;
        })();
      const csv = decodeWindows1256(await readFile((await download.path())!));
      expect(csv.split('\r\n')[0]).toContain(sourceHeaders[14]);
      const outputWorkbook = XLSX.read(csv, { type: 'string', raw: true });
      const outputRows = XLSX.utils.sheet_to_json<Array<string | number | Date>>(outputWorkbook.Sheets[outputWorkbook.SheetNames[0]], {
        header: 1,
        raw: true,
        defval: '',
        blankrows: false,
      });
      expect(outputRows[0]).toEqual(sourceHeaders);
      const outputDates = outputRows.slice(1).map((row) => String(row[10] ?? '').trim()).filter(Boolean);
      expect(outputDates).toEqual(sourceDates);
      const sourceTableGoodsValues = sourceRows.slice(1).map((row) =>
        String(row[1] ?? '').trim() === '1' && !String(row[2] ?? '').trim()
          ? '0'
          : String(row[2] ?? '').trim(),
      );
      const outputTableGoodsValues = outputRows.slice(1).map((row) => String(row[2] ?? '').trim());
      expect(outputTableGoodsValues).toEqual(sourceTableGoodsValues);
      if (sample.direction === 'sales' && selectedDirection === 'sales') {
        expect(outputDates.length).toBeGreaterThan(0);
        expect(outputDates.every((date) => /^\d{2}\/\d{2}\/\d{4}$/.test(date))).toBe(true);
        expect(outputDates).toContain('25/10/2026');
      }
      expect(download.suggestedFilename()).toMatch(new RegExp(`^${sample.direction}_(?:review_incomplete|final_\\d{8})(?:_part_\\d+)?\\.csv$`));
    }
  }
});

test('preserves accepted sales data while asking about blank product fields', async ({ page }) => {
  const workbookPath = resolve(process.cwd(), '../../attached_assets/00_1792956374908.xlsx');
  const acceptedCsvPath = resolve(process.cwd(), '../../attached_assets/00_sales_converted_parts_type3_goods6_legacy_1792990536711.csv');
  test.skip(!existsSync(workbookPath) || !existsSync(acceptedCsvPath), 'The source workbook and accepted CSV are not both present.');

  await uploadWorkbook(page, 'sales', basename(workbookPath), await readFile(workbookPath));
  await expect(page.locator('.missing-values')).toContainText('أسئلة قبل تجهيز الملف النهائي');
  await expect(page.locator('.missing-field').filter({ hasText: 'اسم المنتج' })).toBeVisible();
  await expect(page.locator('.missing-field').filter({ hasText: 'نوع البيان' })).toBeVisible();
  await expect(page.locator('.missing-field').filter({ hasText: 'نوع السلعة' })).toBeVisible();
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  await expect(page.getByRole('alertdialog')).toContainText('غير مكتمل');
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'تنزيل مسودة غير جاهزة للرفع' }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe('sales_review_incomplete.csv');

  const actual = await readFile((await download.path())!);
  const accepted = await readFile(acceptedCsvPath);
  const actualWorkbook = XLSX.read(decodeWindows1256(actual), { type: 'string', raw: true });
  const acceptedWorkbook = XLSX.read(decodeWindows1256(accepted), { type: 'string', raw: true });
  const actualRows = XLSX.utils.sheet_to_json<Array<string | number>>(actualWorkbook.Sheets[actualWorkbook.SheetNames[0]], {
    header: 1,
    raw: true,
    defval: '',
    blankrows: false,
  });
  const acceptedRows = XLSX.utils.sheet_to_json<Array<string | number>>(acceptedWorkbook.Sheets[acceptedWorkbook.SheetNames[0]], {
    header: 1,
    raw: true,
    defval: '',
    blankrows: false,
  });
  expect(actualRows[0]).toEqual(acceptedRows[0]);
  expect(actualRows).toHaveLength(acceptedRows.length);
  const preservedColumns = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 16, 17, 18, 19, 20, 21];
  for (let index = 1; index < actualRows.length; index += 1) {
    expect(preservedColumns.map((column) => String(actualRows[index][column] ?? '')))
      .toEqual(preservedColumns.map((column) => String(acceptedRows[index][column] ?? '')));
    expect(actualRows[index][11]).toBe('');
    expect(actualRows[index][13]).toBe('');
    expect(actualRows[index][14]).toBe('');
  }
  expect(acceptedRows.slice(1).every((row) =>
    row[11] === 'قطع غيار' && String(row[13]) === '3' && String(row[14]) === '6',
  )).toBe(true);
});

test('previews only 20 of 100 rows while keeping the table scrollable on a 360px screen', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 800 });
  const rows = Array.from({ length: 100 }, (_, index) => validRow(index + 1));
  await uploadWorkbook(page, 'sales', 'hundred-rows.xlsx', workbookBuffer(salesHeaders, rows));
  await expect(page.locator('.file-names span')).toContainText('100 صف بيانات');
  await expect(page.locator('.table-scroll tbody tr')).toHaveCount(20);
  await expect(page.locator('.preview-note')).toHaveText('+80 صف إضافي');
  const sizes = await page.evaluate(() => {
    const table = document.querySelector<HTMLElement>('.table-scroll');
    return {
      viewport: window.innerWidth,
      document: document.documentElement.scrollWidth,
      tableViewport: table?.clientWidth ?? 0,
      tableContent: table?.scrollWidth ?? 0,
    };
  });
  expect(sizes.document).toBeLessThanOrEqual(sizes.viewport + 1);
  expect(sizes.tableContent).toBeGreaterThan(sizes.tableViewport);
});

test('splits large CSV exports at row boundaries into files no larger than 3 MB', async ({ page }) => {
  const rows = Array.from({ length: 900 }, (_, index) => {
    const row = validRow(index + 1);
    row[11] = 'x'.repeat(4000);
    return row;
  });
  await uploadWorkbook(page, 'sales', 'large-export.xlsx', workbookBuffer(salesHeaders, rows));
  await expect(page.locator('.report-metrics')).toContainText('2');

  const firstDownloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'إنشاء وتنزيل CSV' }).click();
  const firstDownload = await firstDownloadPromise;
  expect(firstDownload.suggestedFilename()).toMatch(/_part_1\.csv$/);
  const firstBytes = await readFile((await firstDownload.path())!);
  expect(firstBytes.byteLength).toBeLessThanOrEqual(3 * 1024 * 1024);

  const secondDownloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: /تنزيل 2$/ }).click();
  const secondDownload = await secondDownloadPromise;
  expect(secondDownload.suggestedFilename()).toMatch(/_part_2\.csv$/);
  const secondBytes = await readFile((await secondDownload.path())!);
  expect(secondBytes.byteLength).toBeLessThanOrEqual(3 * 1024 * 1024);
});

test('theme and language controls update the page and keep theme after reload', async ({ page }) => {
  const root = page.locator('html');
  await expect(root).toHaveAttribute('data-theme', 'dark');
  await page.getByTestId('button-toggle-theme').click();
  await expect(root).toHaveAttribute('data-theme', 'light');
  await expect(page.getByTestId('button-toggle-theme')).toHaveAttribute('aria-pressed', 'true');
  await page.reload();
  await expect(root).toHaveAttribute('data-theme', 'light');
  await page.getByRole('button', { name: 'Switch language to English' }).click();
  await expect(page.locator('main.app-shell')).toHaveAttribute('dir', 'ltr');
  await expect(page.getByRole('heading', { name: 'Choose invoice type' })).toBeVisible();
  await expect(page.getByTestId('button-toggle-theme')).toHaveAttribute('aria-label', 'Switch to dark mode');
});

test('the installed app shell remains available after the first visit goes offline', async ({ page, context }) => {
  await page.evaluate(async () => {
    if ('serviceWorker' in navigator) await navigator.serviceWorker.ready;
  });
  await page.reload();
  await page.waitForFunction(() => Boolean(navigator.serviceWorker?.controller));
  await context.setOffline(true);
  await page.reload();
  await expect(page.getByRole('heading', { name: /حوّل فواتيرك بسهولة|Convert your tax invoices/ })).toBeVisible();
  await expect(page.getByTestId('button-toggle-theme')).toBeVisible();
  await context.setOffline(false);
});
