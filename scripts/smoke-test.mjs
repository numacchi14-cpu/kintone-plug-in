import { rm } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import ExcelJS from 'exceljs';

const require = createRequire(import.meta.url);

const stamp = `${process.pid}-${Date.now()}`;
const dateRulesOutput = path.join(tmpdir(), `krp-date-rules-${stamp}.mjs`);
const queryOutput = path.join(tmpdir(), `krp-query-${stamp}.mjs`);
const excelOutput = path.join(tmpdir(), `krp-excel-${stamp}.cjs`);

try {
  await Promise.all([
    build({
      entryPoints: ['src/shared/dateRules.ts'],
      bundle: true,
      platform: 'node',
      format: 'esm',
      outfile: dateRulesOutput
    }),
    build({
      entryPoints: ['src/shared/kintoneApi.ts'],
      bundle: true,
      platform: 'node',
      format: 'esm',
      outfile: queryOutput
    }),
    build({
      entryPoints: ['src/shared/excel.ts'],
      bundle: true,
      platform: 'node',
      format: 'cjs',
      outfile: excelOutput
    })
  ]);

  const { calculateDateRange, resolveBaseDate } = await import(pathToFileURL(dateRulesOutput).href);
  const { buildSourceQuery } = await import(pathToFileURL(queryOutput).href);
  const { createInitialTemplate, fillReportTemplate, validateReportTemplate } = require(excelOutput);

  assertEqual(
    resolveBaseDate('yesterday', new Date(2026, 7, 1), ''),
    '2026-07-31',
    '常に昨日の基準日'
  );
  assertEqual(
    resolveBaseDate('yesterday', new Date(2026, 7, 3), ''),
    '2026-08-02',
    '通常日の昨日基準日'
  );
  assertDeepEqual(
    calculateDateRange('baseWeek', '2026-07-15'),
    { start: '2026-07-12', end: '2026-07-18' },
    '基準日の週'
  );
  assertDeepEqual(
    calculateDateRange('baseMonthEnd', '2026-07-15'),
    { start: '2026-07-31', end: '2026-07-31' },
    '基準日の同月末'
  );
  assertDeepEqual(
    calculateDateRange('baseYear', '2026-07-15'),
    { start: '2026-01-01', end: '2026-12-31' },
    '基準日の同年'
  );
  assertDeepEqual(
    calculateDateRange('previousMonthPreviousYear', '2026-08-02'),
    { start: '2025-07-01', end: '2025-07-31' },
    '基準日の前月の前年同月'
  );
  assertDeepEqual(
    calculateDateRange('sameDayPreviousYear', '2024-02-29'),
    { start: '2023-02-28', end: '2023-02-28' },
    'うるう日の前年同日'
  );

  const source = (dateRule) => ({
    key: 'test',
    label: 'テスト',
    appId: '1',
    sheetName: 'テスト',
    fields: [],
    filters: [
      { field: '店舗名', operator: '=', valueFrom: 'store', valueType: 'text' },
      { field: '日付', operator: 'between', valueFrom: 'dateRange', valueType: 'text', dateRule }
    ],
    sorts: [{ field: '日付', order: 'asc' }]
  });

  assertEqual(
    buildSourceQuery(source('monthStartToBaseDate'), '福岡本店', '2026-07-15'),
    '店舗名 = "福岡本店" and 日付 >= "2026-07-01" and 日付 <= "2026-07-15" order by 日付 asc',
    '月次クエリ'
  );
  assertEqual(
    buildSourceQuery(source('yearStartToBaseDate'), '福岡本店', '2026-07-31'),
    '店舗名 = "福岡本店" and 日付 >= "2026-01-01" and 日付 <= "2026-07-31" order by 日付 asc',
    '年次クエリ'
  );
  assertEqual(
    buildSourceQuery(
      {
        ...source('monthStartToBaseDate'),
        filters: [
          { field: '店舗ID', operator: '=', valueFrom: 'outputField', outputField: 'store_id', valueType: 'text' },
          { field: '営業部', operator: '=', valueFrom: 'outputField', outputField: 'department', valueType: 'text' },
          { field: '日付', operator: 'between', valueFrom: 'dateRange', valueType: 'text', dateRule: 'previousMonthPreviousYear' }
        ]
      },
      '',
      '2026-08-02',
      {
        store_id: { type: 'SINGLE_LINE_TEXT', value: 'S001' },
        department: { type: 'SINGLE_LINE_TEXT', value: '西日本' }
      }
    ),
    '店舗ID = "S001" and 営業部 = "西日本" and 日付 >= "2025-07-01" and 日付 <= "2025-07-31" order by 日付 asc',
    '出力アプリ任意フィールド条件クエリ'
  );

  const excelSource = {
    key: 'actual',
    label: '実績',
    appId: '1',
    sheetName: '実績',
    tableName: 'tbl_actual',
    fields: [
      { code: 'date', label: '日付', type: 'date' },
      { code: 'store', label: '店舗名' },
      { code: 'sales', label: '実績_総売上', type: 'number' }
    ],
    filters: [],
    sorts: []
  };
  const templateBuffer = await createInitialTemplate({ sources: [excelSource] }, 'テスト帳票');
  await validateReportTemplate(templateBuffer, [excelSource]);
  const filledBuffer = await fillReportTemplate(
    templateBuffer,
    {
      reportId: 'actual',
      reportName: 'テスト帳票',
      store: '福岡本店',
      baseDate: '2026-07-15',
      periodStart: '2026-07-01',
      periodEnd: '2026-07-15',
      exportedAt: '2026-07-15',
      exporter: 'テスト'
    },
    [
      {
        source: excelSource,
        rows: [
          { date: '2026-07-01', store: '福岡本店', sales: 1000 },
          { date: new Date(2026, 6, 2), store: '福岡本店', sales: 2000 }
        ],
        periodStart: '2026-07-01',
        periodEnd: '2026-07-15',
        query: '店舗名 = "福岡本店"'
      }
    ]
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(filledBuffer));
  const worksheet = workbook.getWorksheet('実績');
  const table = worksheet?.getTable('tbl_actual');
  assertEqual(Boolean(table), true, '元データ用Excelテーブル');
  assertEqual(table.table.tableRef, 'A1:C3', '元データ用Excelテーブル範囲');

  const isoDateCellValue = worksheet.getCell('A2').value;
  assertEqual(isoDateCellValue instanceof Date, true, 'ISO日付文字列をDateへ変換');
  assertEqual(
    isoDateCellValue.getTime(),
    Date.UTC(2026, 6, 1),
    'ISO日付をUTC基準の時刻なしDateとして書き込む（実行環境のタイムゾーンに依存しない）'
  );

  const reportTemplate = new ExcelJS.Workbook();
  await reportTemplate.xlsx.load(Buffer.from(templateBuffer));
  const reportSheet = reportTemplate.addWorksheet('帳票');
  reportSheet.getCell('A1').value = '__PL_FORMULA__:=SUM(1,2)';
  reportSheet.getCell('A1').numFmt = '#,###,;[Red]"△ "#,###,';
  const reportTemplateBuffer = await reportTemplate.xlsx.writeBuffer();
  const reportOutput = await fillReportTemplate(
    reportTemplateBuffer,
    {
      reportId: 'actual',
      reportName: 'テスト帳票',
      store: '福岡本店',
      baseDate: '2026-07-15',
      periodStart: '2026-07-01',
      periodEnd: '2026-07-15',
      exportedAt: '2026-07-15',
      exporter: 'テスト'
    },
    [
      {
        source: excelSource,
        rows: [],
        periodStart: '2026-07-01',
        periodEnd: '2026-07-15',
        query: ''
      }
    ]
  );
  const outputBook = new ExcelJS.Workbook();
  await outputBook.xlsx.load(Buffer.from(reportOutput));
  assertEqual(outputBook.getWorksheet('帳票')?.state, 'visible', '帳票シートを表示');
  assertEqual(outputBook.getWorksheet('設定')?.state, 'hidden', '設定シートを非表示');
  assertEqual(outputBook.getWorksheet('集計')?.state, 'hidden', '集計シートを非表示');
  assertEqual(outputBook.getWorksheet('実績')?.state, 'hidden', '元データ用シートを非表示');
  assertEqual(outputBook.getWorksheet('帳票')?.getCell('A1').value?.formula, 'SUM(1,2)', '数式プレースホルダーを変換');
  assertEqual(outputBook.getWorksheet('帳票')?.getCell('A1').numFmt, '#,###,;[Red]"△ "#,###,', '数式プレースホルダーの表示書式を維持');

  console.log('Smoke tests passed.');
} finally {
  await Promise.all([
    rm(dateRulesOutput, { force: true }),
    rm(queryOutput, { force: true }),
    rm(excelOutput, { force: true })
  ]);
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}\nexpected: ${expected}\nactual:   ${actual}`);
  }
}

function assertDeepEqual(actual, expected, label) {
  assertEqual(JSON.stringify(actual), JSON.stringify(expected), label);
}
