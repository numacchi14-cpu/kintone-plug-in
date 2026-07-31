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

  const { resolveBaseDate } = await import(pathToFileURL(dateRulesOutput).href);
  const { buildSourceQuery } = await import(pathToFileURL(queryOutput).href);
  const { createInitialTemplate, fillReportTemplate, validateReportTemplate } = require(excelOutput);

  assertEqual(
    resolveBaseDate('firstDayUsesYesterday', new Date(2026, 7, 1), ''),
    '2026-07-31',
    '毎月1日の基準日'
  );
  assertEqual(
    resolveBaseDate('today', new Date(2026, 7, 3), '2026-07-31'),
    '2026-07-31',
    '保存済み基準日の優先'
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
          { date: new Date(2026, 6, 1), store: '福岡本店', sales: 1000 },
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
