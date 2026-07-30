import { rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';

const stamp = `${process.pid}-${Date.now()}`;
const dateRulesOutput = path.join(tmpdir(), `krp-date-rules-${stamp}.mjs`);
const queryOutput = path.join(tmpdir(), `krp-query-${stamp}.mjs`);

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
    })
  ]);

  const { resolveBaseDate } = await import(pathToFileURL(dateRulesOutput).href);
  const { buildSourceQuery } = await import(pathToFileURL(queryOutput).href);

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

  console.log('Smoke tests passed.');
} finally {
  await Promise.all([
    rm(dateRulesOutput, { force: true }),
    rm(queryOutput, { force: true })
  ]);
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}\nexpected: ${expected}\nactual:   ${actual}`);
  }
}
