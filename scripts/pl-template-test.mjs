// PL Management（Kintone上の損益管理システム）の実テンプレートを、実際の出力処理（fillReportTemplate）に
// 通して検査する回帰テスト。プラグインを他プロジェクト向けに改修した際に、PL Managementの帳票を
// 壊していないかを確認するためのもの。検査の中身は scripts/lib/template-test-runner.mjs（営業分析の
// scripts/sales-template-test.mjs と共通）。
//
// - テンプレートと取得元アプリ設定JSONはPL Managementのフォルダから読み込み、このリポジトリには
//   コミットしない（業務情報を含むため）。フォルダが無い環境ではスキップする。
// - フォルダの場所は環境変数 PL_MANAGEMENT_DIR で変更できる（既定: C:\Projects\PL Management）。
// - 取得データは実データではなく、取得元アプリ設定JSONのフィールド定義から生成した合成データを使う。
// - --excel を付けると、出力ファイルを実際のExcel（COM）で開いて全再計算し、ファイル破損や
//   #REF!・#NAME?などのエラーセルが無いことも確認する（Windows＋Excelが必要）。
import { runTemplateTests } from './lib/template-test-runner.mjs';

const plRoot = process.env.PL_MANAGEMENT_DIR || 'C:\\Projects\\PL Management';

// 現在Kintoneのテンプレート管理アプリへ登録して運用しているテンプレート。PL Management側で
// テンプレートを差し替えた場合は、この一覧も合わせて更新する。
const cases = [
  {
    name: '単月（2025年度版）',
    template: 'monthly-template-20260812-v0.16.5/単月部門別PL_Kintoneテンプレート_完成版_千円表示.xlsx',
    sources: 'monthly-template-20260812/sources_json_単月部門別PL.json',
    baseDate: '2025-11-30',
    expectedPeriod: ['2025-11-01', '2025-11-30']
  },
  {
    name: '単月（空港店版）',
    template: 'monthly-template-20260812-v0.16.5/単月部門別PL_Kintoneテンプレート_完成版_千円表示_空港店版(2026年度以降用).xlsx',
    sources: 'monthly-template-20260812/sources_json_単月部門別PL.json',
    baseDate: '2026-08-31',
    expectedPeriod: ['2026-08-01', '2026-08-31']
  },
  {
    name: '累計（2025年度版・修正版14）',
    template: 'cumulative-template-20260812/累計部門別PL_Kintoneテンプレート_完成版_千円表示_修正版14.xlsx',
    sources: 'cumulative-template-20260812/sources_json_累計部門別PL.json',
    baseDate: '2025-11-30',
    expectedPeriod: ['2025-01-01', '2025-11-30']
  },
  {
    name: '累計（空港店版）',
    template: 'cumulative-template-20260812/累計部門別PL_Kintoneテンプレート_完成版_千円表示_修正版14_空港店版(2026年度以降用).xlsx',
    sources: 'cumulative-template-20260812/sources_json_累計部門別PL.json',
    baseDate: '2026-08-31',
    expectedPeriod: ['2026-01-01', '2026-08-31']
  },
  {
    name: '時系列（2025年度版）',
    template: 'timeseries-template-20260813/時系列損益計算書_Kintoneテンプレート_完成版_千円表示.xlsx',
    sources: 'timeseries-template-20260813/sources_json_時系列損益計算書.json',
    baseDate: '2025-11-30',
    expectedPeriod: ['2025-01-01', '2025-12-31']
  },
  {
    name: '時系列（空港店版・通年）',
    template: 'timeseries-template-20260813/時系列損益計算書_Kintoneテンプレート_完成版_千円表示_空港店版(2026年度以降用).xlsx',
    sources: 'timeseries-template-20260813/sources_json_時系列損益計算書.json',
    baseDate: '2026-08-15',
    expectedPeriod: ['2026-01-01', '2026-12-31']
  },
  {
    // 基準日が対象期間（1〜6月）の外にあるケース。v0.16.10・v0.16.11の不具合の再現条件。
    name: '時系列（空港店版・上半期）',
    template: 'timeseries-template-20260813/時系列損益計算書_Kintoneテンプレート_完成版_千円表示_空港店版(2026年度以降用)_上半期用.xlsx',
    sources: 'timeseries-template-20260813/sources_json_時系列損益計算書_上半期.json',
    baseDate: '2026-08-15',
    expectedPeriod: ['2026-01-01', '2026-06-30']
  },
  {
    name: '時系列（空港店版・下半期）',
    template: 'timeseries-template-20260813/時系列損益計算書_Kintoneテンプレート_完成版_千円表示_空港店版(2026年度以降用)_下半期用.xlsx',
    sources: 'timeseries-template-20260813/sources_json_時系列損益計算書_下半期.json',
    baseDate: '2026-08-15',
    expectedPeriod: ['2026-07-01', '2026-12-31']
  }
];

await runTemplateTests({
  projectName: 'PL Management',
  projectRoot: plRoot,
  cases,
  rowsFor: (source, range, testCase) => syntheticRows(source, range.start || testCase.expectedPeriod[0], range.end || testCase.expectedPeriod[1]),
  reportId: 'pl_template_test',
  exportedAt: '2026-10-06T09:30:00',
  useExcel: process.argv.includes('--excel')
});

// 取得元アプリ設定JSONのフィールド定義から、各取得元に3行ずつ合成データを作る。開始・終了日に
// あたるフィールド（period_start / valid_from など）には、その取得元の対象期間を入れる。
function syntheticRows(source, start, end) {
  return [0, 1, 2].map((index) =>
    Object.fromEntries(
      source.fields.map((field) => {
        if (/(^|_)(start|from)$/.test(field.code)) return [field.code, start];
        if (/(^|_)(end|to)$/.test(field.code)) return [field.code, end];
        if (field.type === 'number') return [field.code, 1000 * (index + 1)];
        return [field.code, `テスト${index + 1}`];
      })
    )
  );
}
