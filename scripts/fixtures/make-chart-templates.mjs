// グラフ入りテンプレートのテスト用ファイル（scripts/chart-template-test.mjs が使う）を作り直すスクリプト。
// 通常は実行しない（作成済みのファイルをコミットしている）。Windows＋Excel＋Python（openpyxl）が必要。
//
//   node scripts/fixtures/make-chart-templates.mjs
//
// - chart-base.xlsx: createInitialTemplate で作った初期テンプレートに、集計シートのグラフ用の範囲
//   （数式プレースホルダー）と帳票シートを加えたもの。業務データは含まない。
// - chart-excel.xlsx: chart-base.xlsx をExcelで開き、帳票シートに折れ線グラフを2つ置いて保存したもの
//   （相対パスの関係、chartStyle・colorsの部品つき）。
// - chart-openpyxl.xlsx: chart-base.xlsx をopenpyxlで開き、帳票シートに折れ線グラフを置いて保存したもの
//   （描画・グラフへの関係が絶対パス /xl/...）。テーブルへの参照だけは、営業分析の build_template.py と
//   同じく相対パスへ直す（ExcelJSが絶対パスのテーブル参照を読めないため）。
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const require = createRequire(import.meta.url);
const fixtureDir = path.dirname(fileURLToPath(import.meta.url));
const workDir = await mkdtemp(path.join(tmpdir(), 'krp-chart-fixtures-'));

const chartSource = {
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

try {
  const excelModule = path.join(workDir, 'excel.cjs');
  await build({ entryPoints: ['src/shared/excel.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: excelModule });
  const { createInitialTemplate } = require(excelModule);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await createInitialTemplate({ sources: [chartSource] }, 'グラフテスト'));
  const summary = workbook.getWorksheet('集計');
  summary.getCell('A1').value = 'グラフ用';
  summary.getCell('A3').value = '売上';
  summary.getCell('A4').value = '件数';
  for (let month = 1; month <= 12; month++) {
    summary.getCell(2, month + 1).value = `${month}月`;
    const range = `DATE(2026,${month},1),tbl_actual[日付],"<"&DATE(2026,${month + 1},1)`;
    summary.getCell(3, month + 1).value = `__PL_FORMULA__:=SUMIFS(tbl_actual[実績_総売上],tbl_actual[日付],">="&${range})`;
    summary.getCell(4, month + 1).value = `__PL_FORMULA__:=COUNTIFS(tbl_actual[日付],">="&${range})`;
  }
  const report = workbook.addWorksheet('帳票');
  report.getCell('A1').value = 'グラフテスト';
  report.getCell('A2').value = '__PL_FORMULA__:=SUM(集計!B3:M3)';
  report.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
  const basePath = path.join(fixtureDir, 'chart-base.xlsx');
  await writeFile(basePath, Buffer.from(await workbook.xlsx.writeBuffer()));

  // Excel: 折れ線グラフ（AddChart2、スタイル227）を2つ置く。
  const excelPath = path.join(fixtureDir, 'chart-excel.xlsx');
  const script = path.join(workDir, 'excel.ps1');
  // Windows PowerShell 5.1はBOMなしのUTF-8をANSIとして読むため、シート名の日本語のためにBOMを付ける。
  await writeFile(
    script,
    `\ufeff$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $workbook = $excel.Workbooks.Open('${basePath}')
  $sheet = $workbook.Worksheets.Item('帳票')
  $summary = $workbook.Worksheets.Item('集計')
  $first = $sheet.Shapes.AddChart2(227, 4, 10, 60, 420, 220).Chart
  $first.SetSourceData($summary.Range('A2:M3'))
  $first.HasTitle = $true
  $first.ChartTitle.Text = '売上'
  $second = $sheet.Shapes.AddChart2(227, 4, 10, 300, 420, 220).Chart
  $second.SetSourceData($summary.Range('A2:M2,A4:M4'))
  $workbook.SaveAs('${excelPath}', 51)
  $workbook.Close($false)
} finally {
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
`,
    'utf8'
  );
  execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script], { stdio: 'inherit' });

  // openpyxl: 折れ線グラフを1つ置く。
  const openpyxlPath = path.join(fixtureDir, 'chart-openpyxl.xlsx');
  const python = path.join(workDir, 'openpyxl_chart.py');
  await writeFile(
    python,
    `import sys
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
wb = load_workbook(sys.argv[1])
summary = wb['集計']
chart = LineChart()
chart.title = '売上'
chart.add_data(Reference(summary, min_col=1, max_col=13, min_row=3, max_row=4), from_rows=True, titles_from_data=True)
chart.set_categories(Reference(summary, min_col=2, max_col=13, min_row=2, max_row=2))
wb['帳票'].add_chart(chart, 'B5')
wb.save(sys.argv[2])
`,
    'utf8'
  );
  execFileSync('python', ['-I', python, basePath, openpyxlPath], { stdio: 'inherit' });
  const openpyxlZip = await JSZip.loadAsync(await readFile(openpyxlPath));
  for (const name of Object.keys(openpyxlZip.files).filter((file) => /^xl\/worksheets\/_rels\/.+\.rels$/.test(file))) {
    const xml = await openpyxlZip.file(name).async('string');
    openpyxlZip.file(name, xml.replace(/Target="\/xl\/tables\//g, 'Target="../tables/'));
  }
  await writeFile(openpyxlPath, await openpyxlZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));

  // 公開リポジトリに置くため、作成者などの個人情報を消す。
  for (const file of [basePath, excelPath, openpyxlPath]) {
    const zip = await JSZip.loadAsync(await readFile(file));
    const core = zip.file('docProps/core.xml');
    if (core) {
      const xml = await core.async('string');
      zip.file('docProps/core.xml', xml.replace(/<(dc:creator|cp:lastModifiedBy)>[^<]*<\/\1>/g, '<$1></$1>'));
    }
    const app = zip.file('docProps/app.xml');
    if (app) {
      const xml = await app.async('string');
      zip.file('docProps/app.xml', xml.replace(/<(Company|Manager)>[^<]*<\/\1>/g, '<$1></$1>'));
    }
    await writeFile(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  }
  console.log('chart fixtures written to', fixtureDir);
} finally {
  await rm(workDir, { recursive: true, force: true });
}
