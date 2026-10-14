// 営業分析（C:\Projects\営業分析）の実テンプレートを、実際の出力処理（fillReportTemplate）に通して検査する
// 回帰テスト。本番運用に入った営業分析の帳票を、プラグインの改修から守るためのもの（2026-10-14追加）。
// 検査の中身は scripts/lib/template-test-runner.mjs（PL Managementの scripts/pl-template-test.mjs と共通）。
//
// - テンプレートと取得元アプリ設定JSONは営業分析のフォルダから読み込み、このリポジトリにはコミットしない
//   （業務情報を含むため）。フォルダが無い環境ではスキップする。
// - フォルダの場所は環境変数 SALES_ANALYSIS_DIR で変更できる（既定: C:\Projects\営業分析）。
// - 取得データは実データではなく、Kintone REST APIと同じ形（値はすべて文字列）の合成データを作り、
//   実際の出力と同じく fieldDisplayValue を通して渡す。店舗名・イベント名は架空のもの。
// - 数値の正しさ（旧資料・仕様どおりの集計か）は確認しない。それは営業分析側の verify/（全セル照合）で行う。
// - --excel を付けると、出力ファイルを実際のExcel（COM）で開いて全再計算し、ファイル破損やエラーセルが無いこと、
//   帳票シートのグラフ数も確認する（Windows＋Excelが必要）。
import { runTemplateTests } from './lib/template-test-runner.mjs';

const salesRoot = process.env.SALES_ANALYSIS_DIR || 'C:\\Projects\\営業分析';
const store = 'テスト店';

// 現在Kintoneのテンプレート管理アプリ（営業分析）へ登録して運用しているテンプレート。営業分析側で
// テンプレートを差し替えた場合は、この一覧も合わせて更新する。
const cases = [
  {
    name: '日平均（グラフ入り）',
    template: '日平均-template/日平均_完成版テンプレート.xlsx',
    sources: '日平均-template/sources.json',
    store,
    baseDate: '2026-10-13',
    expectedPeriod: ['2025-01-01', '2026-10-13'],
    charts: 2
  },
  {
    // 当年が1月だけのケース。当年1月の前月対比は前年12月と比べ、グラフ用データの2〜12月は値なしになる。
    name: '日平均（グラフ入り・基準日が1月）',
    template: '日平均-template/日平均_完成版テンプレート.xlsx',
    sources: '日平均-template/sources.json',
    store,
    baseDate: '2026-01-05',
    expectedPeriod: ['2025-01-01', '2026-01-05'],
    charts: 2
  },
  {
    name: '日平均（グラフなし）',
    template: '日平均-template/日平均_完成版テンプレート_グラフなし.xlsx',
    sources: '日平均-template/sources.json',
    store,
    baseDate: '2026-10-13',
    expectedPeriod: ['2025-01-01', '2026-10-13'],
    charts: 0
  }
];

// イベント日設定（1レコード＝ルール1つ）。前年の途中で終わる枠、当年に始まる枠、1つの枠に複数の種類を含める。
const eventRules = [
  ['1', '7の日', '末尾', '7', '2025-01-01', ''],
  ['2', '5の日', '末尾', '5', '2025-01-01', '2025-06-30'],
  ['3', '土曜', '曜日', '土', '2026-04-01', ''],
  ['4', '15日・特定日', '日', '15', '2025-01-01', ''],
  ['4', '15日・特定日', '日付指定', '', '2025-08-10', '2025-08-12']
];

await runTemplateTests({
  projectName: '営業分析',
  projectRoot: salesRoot,
  cases,
  rowsFor: (source, range, testCase, { fieldDisplayValue }) =>
    kintoneRecords(source, range, testCase).map((record) =>
      Object.fromEntries(source.fields.map((field) => [field.code, fieldDisplayValue(record, field)]))
    ),
  reportId: 'sales_template_test',
  exportedAt: '2026-10-14T09:30:00',
  useExcel: process.argv.includes('--excel')
});

// Kintone REST APIと同じ形（{ フィールドコード: { type, value } }、値は文字列）の合成レコードを作る。
function kintoneRecords(source, range, testCase) {
  const fieldTypes = Object.fromEntries(source.fields.map((field) => [field.code, field.type === 'number' ? 'NUMBER' : field.type === 'date' ? 'DATE' : 'SINGLE_LINE_TEXT']));
  const record = (values) => Object.fromEntries(Object.entries(values).map(([code, value]) => [code, { type: fieldTypes[code] ?? 'SINGLE_LINE_TEXT', value }]));

  if (!range.start) {
    return eventRules.map(([order, name, kind, value, from, to]) =>
      record({ 店舗名: testCase.store, 表示順: order, イベント名: name, 種類: kind, 値: value, 適用開始日: from, 適用終了日: to })
    );
  }

  // 見通し: 対象期間の毎日、種類ごとに1行。月の1日は休業（営業日数0・売上0）にして、営業日0の扱いも通す。
  const records = [];
  for (let time = Date.parse(`${range.start}T00:00:00Z`); time <= Date.parse(`${range.end}T00:00:00Z`); time += 86400000) {
    const date = new Date(time);
    const day = date.toISOString().slice(0, 10);
    const open = date.getUTCDate() !== 1;
    const holiday = date.getUTCDay() === 0 || date.getUTCDay() === 6;
    [['パチンコ', '4', '機種A', 1], ['スロット', '20', '機種B', 2]].forEach(([kind, rate, model, factor]) => {
      const sales = open ? (300000 + ((time / 86400000) % 31) * 7000) * factor : 0;
      records.push(
        record({
          日付: day,
          店舗名: testCase.store,
          種類: kind,
          レート: rate,
          型: model,
          平日休日区分: holiday ? '休日' : '平日',
          実績_営業日数: open ? '1' : '0',
          実績_総売上: String(sales),
          実績_総粗利: String(Math.round(sales * 0.12))
        })
      );
    });
  }
  return records;
}
