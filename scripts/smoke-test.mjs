import { rm } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const require = createRequire(import.meta.url);

const stamp = `${process.pid}-${Date.now()}`;
const dateRulesOutput = path.join(tmpdir(), `krp-date-rules-${stamp}.mjs`);
const queryOutput = path.join(tmpdir(), `krp-query-${stamp}.mjs`);
const excelOutput = path.join(tmpdir(), `krp-excel-${stamp}.cjs`);
const configOutput = path.join(tmpdir(), `krp-config-${stamp}.mjs`);

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
    }),
    build({
      entryPoints: ['src/shared/config.ts'],
      bundle: true,
      platform: 'node',
      format: 'esm',
      outfile: configOutput
    })
  ]);

  const { calculateDateRange, resolveBaseDate } = await import(pathToFileURL(dateRulesOutput).href);
  const { buildSourceQuery, sourceDateRange, mergeSourceRanges } = await import(pathToFileURL(queryOutput).href);
  const { createInitialTemplate, fillReportTemplate, validateReportTemplate, fixSheetPrElementOrder } = require(excelOutput);
  const { parseTemplateSources } = await import(pathToFileURL(configOutput).href);

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
    calculateDateRange('baseFirstHalf', '2026-07-15'),
    { start: '2026-01-01', end: '2026-06-30' },
    '基準日の同年上半期'
  );
  assertDeepEqual(
    calculateDateRange('baseSecondHalf', '2026-07-15'),
    { start: '2026-07-01', end: '2026-12-31' },
    '基準日の同年下半期'
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

  // 2026-08-15発見の不具合の回帰テスト：日付範囲フィルターを持たないソース（配賦設定履歴など）は
  // 期間開始・終了ともに空文字列を返し、基準日にフォールバックしてはならない。フォールバックすると、
  // mergeSourceRangesで他の日付フィルター付きソースの本当の期間ではなく、この空ダミー値（基準日）が
  // 誤って採用されてしまう（特に基準日が対象期間の外にある上半期・下半期出力で顕在化した）。
  assertDeepEqual(
    sourceDateRange(source('baseFirstHalf'), '2025-12-15'),
    { start: '2025-01-01', end: '2025-06-30' },
    '日付フィルターありソースの期間（上半期）'
  );
  assertDeepEqual(
    sourceDateRange({ ...source('baseFirstHalf'), filters: [] }, '2025-12-15'),
    { start: '', end: '' },
    '日付フィルターなしソースは空文字列（基準日にフォールバックしない）'
  );
  assertDeepEqual(
    mergeSourceRanges(
      [
        { periodStart: '2025-01-01', periodEnd: '2025-06-30' }, // PL実績明細（上半期フィルターあり）
        { periodStart: '', periodEnd: '' } // 配賦設定履歴（フィルターなし）
      ],
      '2025-12-15'
    ),
    { periodStart: '2025-01-01', periodEnd: '2025-06-30' },
    '基準日が対象期間の外でも、フィルターなしソースの空値に引きずられず正しい期間を採用する'
  );
  assertDeepEqual(
    mergeSourceRanges([{ periodStart: '', periodEnd: '' }], '2025-12-15'),
    { periodStart: '2025-12-15', periodEnd: '2025-12-15' },
    '全ソースが日付フィルターなしの場合は基準日にフォールバックする'
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

  // ExcelJSが<sheetPr>の子要素順序（正しくはtabColor, outlinePr, pageSetUpPrの順）を崩して
  // 書き出すことがあり、OOXMLスキーマ順序違反となってExcelがシート内容ごと読み込み拒否・破棄する
  // 不具合の回帰テスト（2026-08-15発見・修正）。
  const brokenZip = new JSZip();
  brokenZip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetPr><pageSetUpPr fitToPage="1"/><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><sheetData/></worksheet>'
  );
  const brokenBuffer = await brokenZip.generateAsync({ type: 'arraybuffer' });
  const fixedBuffer = await fixSheetPrElementOrder(brokenBuffer);
  const fixedZip = await JSZip.loadAsync(fixedBuffer);
  const fixedXml = await fixedZip.file('xl/worksheets/sheet1.xml').async('string');
  assertEqual(
    fixedXml.includes('<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr>'),
    true,
    'sheetPrの子要素順序をOOXMLスキーマどおり（outlinePr→pageSetUpPr）に補正'
  );

  // 元から正しい順序の場合は変更しない（他の属性・空白等を壊さないことの確認）
  const okZip = new JSZip();
  okZip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>'
  );
  const okBuffer = await okZip.generateAsync({ type: 'arraybuffer' });
  const okResultBuffer = await fixSheetPrElementOrder(okBuffer);
  const okResultZip = await JSZip.loadAsync(okResultBuffer);
  const okResultXml = await okResultZip.file('xl/worksheets/sheet1.xml').async('string');
  assertEqual(
    okResultXml,
    '<worksheet><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>',
    '既に正しい順序のsheetPrは変更しない'
  );

  // 2026-08-15発見の不具合の回帰テスト：型定義（DateRangeRule）へ新しい日付ルールを追加しても、
  // parseTemplateSourcesが使う実行時の許可リスト（dateRangeRules配列）を同時に更新し忘れると、
  // Kintoneに保存された取得元アプリ設定JSONをパースする際にそのルールが黙って'sameDay'へ差し替え
  // られてしまう（型チェックでは検出できず、直接オブジェクトを渡すユニットテストでも素通りしてしまう
  // ため発覚が遅れた。実際にJSON文字列としてパースする経路でのみ再現する）。
  const sourcesJsonWithHalfYearRule = JSON.stringify([
    {
      key: 'pl_actuals',
      label: 'PL実績明細',
      appId: '175',
      sheetName: 'PL実績明細',
      fields: [{ code: 'period_start', label: 'period_start', type: 'text' }],
      filters: [{ field: 'period_start', operator: 'between', valueFrom: 'dateRange', valueType: 'text', dateRule: 'baseFirstHalf' }],
      sorts: []
    }
  ]);
  const parsedSources = parseTemplateSources(sourcesJsonWithHalfYearRule, []);
  assertEqual(
    parsedSources[0].filters[0].dateRule,
    'baseFirstHalf',
    'JSON文字列からのパースでも新しい日付ルール（上半期）が保持される（sameDayへ差し替わらない）'
  );

  console.log('Smoke tests passed.');
} finally {
  await Promise.all([
    rm(dateRulesOutput, { force: true }),
    rm(queryOutput, { force: true }),
    rm(excelOutput, { force: true }),
    rm(configOutput, { force: true })
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
