// PL Management（Kintone上の損益管理システム）の実テンプレートを、実際の出力処理（fillReportTemplate）に
// 通して検査する回帰テスト。プラグインを他プロジェクト向けに改修した際に、PL Managementの帳票を
// 壊していないかを確認するためのもの。
//
// - テンプレートと取得元アプリ設定JSONはPL Managementのフォルダから読み込み、このリポジトリには
//   コミットしない（業務情報を含むため）。フォルダが無い環境ではスキップする。
// - フォルダの場所は環境変数 PL_MANAGEMENT_DIR で変更できる（既定: C:\Projects\PL Management）。
// - 取得データは実データではなく、取得元アプリ設定JSONのフィールド定義から生成した合成データを使う。
// - --excel を付けると、出力ファイルを実際のExcel（COM）で開いて全再計算し、ファイル破損や
//   #REF!・#NAME?などのエラーセルが無いことも確認する（Windows＋Excelが必要）。
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const require = createRequire(import.meta.url);

const plRoot = process.env.PL_MANAGEMENT_DIR || 'C:\\Projects\\PL Management';
const useExcel = process.argv.includes('--excel');
const formulaPlaceholderPrefix = '__PL_FORMULA__:';
const settingsSheetName = '設定';
const summarySheetName = '集計';

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

if (!existsSync(plRoot)) {
  console.log(`PL Managementのフォルダが見つからないため、実テンプレートテストをスキップしました: ${plRoot}`);
  process.exit(0);
}

const stamp = `${process.pid}-${Date.now()}`;
const queryOutput = path.join(tmpdir(), `krp-pl-query-${stamp}.mjs`);
const excelOutput = path.join(tmpdir(), `krp-pl-excel-${stamp}.cjs`);
const configOutput = path.join(tmpdir(), `krp-pl-config-${stamp}.mjs`);
const failures = [];

try {
  await Promise.all([
    build({ entryPoints: ['src/shared/kintoneApi.ts'], bundle: true, platform: 'node', format: 'esm', outfile: queryOutput }),
    build({ entryPoints: ['src/shared/excel.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: excelOutput }),
    build({ entryPoints: ['src/shared/config.ts'], bundle: true, platform: 'node', format: 'esm', outfile: configOutput })
  ]);

  const { buildSourceQuery, sourceDateRange, mergeSourceRanges } = await import(pathToFileURL(queryOutput).href);
  const { fillReportTemplate, validateReportTemplate } = require(excelOutput);
  const { parseTemplateSources } = await import(pathToFileURL(configOutput).href);

  const outputDir = useExcel ? await mkdtemp(path.join(tmpdir(), 'krp-pl-output-')) : '';
  const excelTargets = [];

  for (const testCase of cases) {
    let caseFailures = 0;
    const fail = (message) => {
      caseFailures += 1;
      if (caseFailures <= 20) failures.push(`[${testCase.name}] ${message}`);
      else if (caseFailures === 21) failures.push(`[${testCase.name}] ……以降の失敗は省略`);
    };
    const templatePath = path.join(plRoot, 'outputs', testCase.template);
    const sourcesPath = path.join(plRoot, 'outputs', testCase.sources);
    if (!existsSync(templatePath) || !existsSync(sourcesPath)) {
      fail(`テンプレートまたは取得元アプリ設定JSONが見つかりません。PL Management側で差し替えた場合は、このテストの一覧を更新してください。\n  ${templatePath}\n  ${sourcesPath}`);
      continue;
    }

    // Kintoneに保存された取得元アプリ設定JSON（文字列）をパースする、実際の入力経路を通す。
    const sources = parseTemplateSources(await readFile(sourcesPath, 'utf8'), []);
    const templateBuffer = toArrayBuffer(await readFile(templatePath));
    await validateReportTemplate(templateBuffer, sources);

    const sourceRows = sources.map((source) => {
      const range = sourceDateRange(source, testCase.baseDate);
      return {
        source,
        rows: syntheticRows(source, range.start || testCase.expectedPeriod[0], range.end || testCase.expectedPeriod[1]),
        periodStart: range.start,
        periodEnd: range.end,
        query: buildSourceQuery(source, '', testCase.baseDate)
      };
    });
    const wholeRange = mergeSourceRanges(sourceRows, testCase.baseDate);
    const context = {
      reportId: 'pl_template_test',
      reportName: testCase.name,
      store: '',
      baseDate: testCase.baseDate,
      periodStart: wholeRange.periodStart,
      periodEnd: wholeRange.periodEnd,
      exportedAt: '2026-10-06T09:30:00',
      exporter: 'テスト'
    };
    if (context.periodStart !== testCase.expectedPeriod[0] || context.periodEnd !== testCase.expectedPeriod[1]) {
      fail(`対象期間が想定と違います。expected: ${testCase.expectedPeriod.join('〜')} actual: ${context.periodStart}〜${context.periodEnd}`);
    }

    const outputBuffer = await fillReportTemplate(templateBuffer, context, sourceRows);

    const template = new ExcelJS.Workbook();
    await template.xlsx.load(templateBuffer);
    const output = new ExcelJS.Workbook();
    await output.xlsx.load(outputBuffer);

    checkFormulas(template, output, fail);
    checkSettingsSheet(template, output, context, fail);
    checkSheetVisibility(template, output, sourceRows, fail);
    checkSourceTables(output, sourceRows, fail);
    await checkWorkbookXml(outputBuffer, fail);

    if (useExcel) {
      const outputPath = path.join(outputDir, `case${excelTargets.length + 1}.xlsx`);
      await writeFile(outputPath, Buffer.from(outputBuffer));
      excelTargets.push({ name: testCase.name, path: outputPath });
    }
    console.log(`  ${testCase.name}: ExcelJSでの検査が完了`);
  }

  if (useExcel && excelTargets.length) {
    await checkWithExcel(excelTargets, outputDir, failures);
  }
  if (outputDir) {
    await rm(outputDir, { recursive: true, force: true });
  }

  if (failures.length) {
    throw new Error(['PL Management実テンプレートテストに失敗しました。', ...failures].join('\n'));
  }
  console.log(`PL Management template tests passed (${cases.length} templates${useExcel ? ', Excel verified' : ''}).`);
} finally {
  await Promise.all([
    rm(queryOutput, { force: true }),
    rm(excelOutput, { force: true }),
    rm(configOutput, { force: true })
  ]);
}

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

// 数式プレースホルダーがすべて、テンプレートどおりの通常の数式に変換されていること。
function checkFormulas(template, output, fail) {
  let placeholderCount = 0;
  template.eachSheet((worksheet) => {
    const outputSheet = output.getWorksheet(worksheet.name);
    eachCell(worksheet, (cell) => {
      const expected = placeholderFormula(cell);
      if (expected === null) return;
      placeholderCount += 1;
      const actual = outputSheet?.getCell(cell.address).formula;
      if (actual !== expected) {
        fail(`${worksheet.name}!${cell.address} の数式プレースホルダーが変換されていません。expected: ${truncate(expected)} actual: ${truncate(String(actual))}`);
      }
    });
  });
  if (!placeholderCount) {
    fail('テンプレートに数式プレースホルダーが1件もありません。テンプレートの差し替えを確認してください。');
  }

  output.eachSheet((worksheet) => {
    eachCell(worksheet, (cell) => {
      if (placeholderFormula(cell) !== null) {
        fail(`${worksheet.name}!${cell.address} に数式プレースホルダーが残っています。`);
      }
      if (cell.formula && cell.formula.length > 8192) {
        fail(`${worksheet.name}!${cell.address} の数式がExcelの上限8,192文字を超えています（${cell.formula.length}文字）。`);
      }
    });
  });
}

// 設定シートの互換性の約束（SPEC.md「設定シートの互換性の約束」）を、実テンプレートで確認する。
// - B4〜B6はISO形式の文字列、B7はExcel日付で書き込まれる
// - テンプレートが置いた設定シートのセル（A1:B8以外）は、出力後もそのまま残る
// - テンプレートの数式が参照する設定シートのセルは、B1:B8か、テンプレート自身が置いたセルだけ
//   （プラグインが出力のたびに上書きする取得元メタデータを参照していない）
function checkSettingsSheet(template, output, context, fail) {
  const templateSheet = template.getWorksheet(settingsSheetName);
  const outputSheet = output.getWorksheet(settingsSheetName);

  [
    ['B4', context.baseDate],
    ['B5', context.periodStart],
    ['B6', context.periodEnd]
  ].forEach(([address, expected]) => {
    const value = outputSheet.getCell(address).value;
    if (value !== expected) {
      fail(`設定!${address} はISO形式の文字列「${expected}」であるべきですが、${JSON.stringify(value)} でした。`);
    }
  });
  if (!(outputSheet.getCell('B7').value instanceof Date)) {
    fail(`設定!B7（出力日）はExcel日付であるべきですが、${JSON.stringify(outputSheet.getCell('B7').value)} でした。`);
  }

  const templateOwned = new Set();
  eachCell(templateSheet, (cell) => {
    if (cell.row <= 8 && cell.col <= 2) return;
    templateOwned.add(cell.address);
    const expected = normalizedContent(cell);
    const actual = normalizedContent(outputSheet.getCell(cell.address));
    if (expected !== actual) {
      fail(`テンプレートが置いた設定!${cell.address} が出力時に上書き・消去されました。expected: ${truncate(expected)} actual: ${truncate(actual)}`);
    }
  });

  output.eachSheet((worksheet) => {
    eachCell(worksheet, (cell) => {
      if (!cell.formula) return;
      for (const address of settingsReferences(cell.formula, worksheet.name === settingsSheetName)) {
        const { row, col } = splitAddress(address);
        const pluginWritten = row <= 8 && col === 2;
        if (!pluginWritten && !templateOwned.has(address)) {
          fail(`${worksheet.name}!${cell.address} の数式が、プラグインの約束の外にある設定!${address} を参照しています（取得元メタデータなど、出力のたびに書き換わるセルの可能性があります）。`);
        }
      }
    });
  });
}

// 技術用シート（設定・集計・元データ）は非表示、帳票シートはテンプレートどおり表示されること。
function checkSheetVisibility(template, output, sourceRows, fail) {
  const technical = new Set([settingsSheetName, summarySheetName, ...sourceRows.map(({ source }) => source.sheetName)]);
  output.eachSheet((worksheet) => {
    if (technical.has(worksheet.name)) {
      if (worksheet.state !== 'hidden') fail(`技術用シート「${worksheet.name}」が非表示になっていません。`);
      return;
    }
    const templateState = template.getWorksheet(worksheet.name)?.state;
    if (templateState === 'visible' && worksheet.state !== 'visible') {
      fail(`帳票シート「${worksheet.name}」が表示されていません。`);
    }
  });
  const activeSheet = output.worksheets[output.views?.[0]?.activeTab ?? -1];
  if (!activeSheet || technical.has(activeSheet.name)) {
    fail(`アクティブなシートが帳票シートではありません（${activeSheet?.name ?? 'なし'}）。`);
  }
}

// 元データ用シートが、取得件数どおりの範囲のExcelテーブルになっていること。
function checkSourceTables(output, sourceRows, fail) {
  sourceRows.forEach(({ source, rows }) => {
    const worksheet = output.getWorksheet(source.sheetName);
    const table = worksheet?.getTables().find(({ table }) => table.tableRef?.startsWith('A1:'));
    const expectedRef = `A1:${columnLetter(source.fields.length)}${Math.max(rows.length, 1) + 1}`;
    if (!table) {
      fail(`元データ用シート「${source.sheetName}」にExcelテーブルがありません。`);
    } else if (table.table.tableRef !== expectedRef) {
      fail(`元データ用シート「${source.sheetName}」のテーブル範囲が想定と違います。expected: ${expectedRef} actual: ${table.table.tableRef}`);
    }
  });
}

// Excelで開いたときの全再計算指定と、<sheetPr>の子要素順序（v0.16.9のファイル破損対策）。
async function checkWorkbookXml(outputBuffer, fail) {
  const zip = await JSZip.loadAsync(outputBuffer);
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  if (!/<calcPr\b[^>]*fullCalcOnLoad="1"/.test(workbookXml)) {
    fail('ブックにExcel起動時の全再計算指定（fullCalcOnLoad）がありません。');
  }
  const order = ['tabColor', 'outlinePr', 'pageSetUpPr'];
  for (const name of Object.keys(zip.files).filter((file) => /^xl\/worksheets\/sheet\d+\.xml$/.test(file))) {
    const xml = await zip.file(name).async('string');
    const sheetPr = xml.match(/<sheetPr(?:\s[^>]*)?>([\s\S]*?)<\/sheetPr>/);
    if (!sheetPr) continue;
    const children = [...sheetPr[1].matchAll(/<(tabColor|outlinePr|pageSetUpPr)\b/g)].map((match) => order.indexOf(match[1]));
    if (children.some((value, index) => index > 0 && value < children[index - 1])) {
      fail(`${name} の<sheetPr>の子要素順序がOOXMLスキーマ違反です（Excelでファイルが壊れます）。`);
    }
  }
}

// 出力ファイルを実際のExcelで開き、全再計算して、表示中の帳票シートのエラーセルを数える。
// Excelが修復を必要とするファイル（v0.16.9のような破損）は、自動化では開けずに失敗する。
async function checkWithExcel(targets, outputDir, failures) {
  const scriptPath = path.join(outputDir, 'check.ps1');
  const listPath = path.join(outputDir, 'targets.json');
  await writeFile(listPath, JSON.stringify(targets), 'utf8');
  await writeFile(
    scriptPath,
    `$ErrorActionPreference = 'Stop'
$targets = Get-Content -LiteralPath '${listPath}' -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$results = @()
try {
  foreach ($target in $targets) {
    $result = [ordered]@{ name = $target.name; opened = $false; formulaCount = 0; errors = @{}; samples = @() }
    try {
      $workbook = $excel.Workbooks.Open($target.path, 0, $true)
      $result.opened = $true
      $excel.CalculateFullRebuild()
      foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Visible -ne -1) { continue }
        try { $result.formulaCount += $sheet.UsedRange.SpecialCells(-4123).Count } catch {}
        try { $errorCells = $sheet.UsedRange.SpecialCells(-4123, 16) } catch { continue }
        foreach ($cell in $errorCells.Cells) {
          $text = [string]$cell.Text
          if ($result.errors.ContainsKey($text)) { $result.errors[$text] += 1 } else { $result.errors[$text] = 1 }
          if ($text -ne '#DIV/0!' -and $result.samples.Count -lt 5) { $result.samples += ($sheet.Name + '!' + $cell.Address($false, $false) + '=' + $text) }
        }
      }
      $workbook.Close($false)
    } catch {
      $result.openError = $_.Exception.Message
    }
    $results += [pscustomobject]$result
  }
} finally {
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
$results | ConvertTo-Json -Depth 5 -Compress
`,
    'utf8'
  );

  console.log('  Excelで出力ファイルを開いて全再計算しています...');
  const stdout = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath], {
    encoding: 'utf8',
    maxBuffer: 16 * 1024 * 1024
  });
  // #DIV/0!は、売上0の店舗の比率など取得データ次第で旧帳票でも出るため失敗にしない。#REF!・#NAME?・
  // #VALUE!・#N/Aなどは、参照切れや設定シートの型の不一致（例：B5/B6が日付になりDATEVALUEが失敗）を示す。
  const dataDependentErrors = new Set(['#DIV/0!']);
  const results = [JSON.parse(stdout.trim())].flat();
  for (const result of results) {
    const errors = Object.entries(result.errors ?? {});
    const breakingErrors = errors.filter(([text]) => !dataDependentErrors.has(text));
    const note = errors.length ? `、データ依存のエラー ${JSON.stringify(Object.fromEntries(errors))}` : '';
    if (!result.opened) {
      failures.push(`[${result.name}] Excelで出力ファイルを開けません（ファイル破損の可能性）: ${result.openError}`);
    } else if (!result.formulaCount) {
      failures.push(`[${result.name}] Excelで開いた帳票シートに数式が1件もありません（シート内容が破棄された可能性）。`);
    } else if (breakingErrors.length) {
      failures.push(`[${result.name}] Excelの全再計算後にエラーセルがあります: ${JSON.stringify(Object.fromEntries(breakingErrors))} 例: ${result.samples.join(', ')}`);
    } else {
      console.log(`  ${result.name}: Excelで開けて全再計算できた（数式${result.formulaCount}件${note}）`);
    }
  }
}

// 結合セルは左上のセルだけを対象にする（ExcelJSは結合範囲の他のセルにも左上の値を複製して見せる）。
function eachCell(worksheet, callback) {
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.isMerged && cell.master !== cell) return;
      callback(cell);
    });
  });
}

// テンプレートのセルが数式プレースホルダーなら、変換後に期待される数式を返す（excel.tsと同じ規則）。
function placeholderFormula(cell) {
  if (typeof cell.value === 'string' && cell.value.startsWith(formulaPlaceholderPrefix)) {
    return cell.value.slice(formulaPlaceholderPrefix.length).replace(/^=/, '');
  }
  if (typeof cell.result === 'string' && cell.result.startsWith(formulaPlaceholderPrefix) && cell.formula) {
    return cell.formula;
  }
  return null;
}

function normalizedContent(cell) {
  const formula = placeholderFormula(cell) ?? cell.formula;
  if (formula) return `=${formula}`;
  const value = cell.value;
  return value instanceof Date ? value.toISOString() : JSON.stringify(value ?? null);
}

// 数式中の設定シートへの参照（範囲は展開）を列挙する。設定シート上の数式は、シート名なしの参照も対象。
function settingsReferences(formula, onSettingsSheet) {
  const addresses = [];
  const cellPattern = '\\$?([A-Z]{1,3})\\$?(\\d+)(?::\\$?([A-Z]{1,3})\\$?(\\d+))?';
  const patterns = [new RegExp(`'?${settingsSheetName}'?!${cellPattern}`, 'g')];
  if (onSettingsSheet) {
    patterns.push(new RegExp(`(?<![!A-Za-z0-9_'"\\]])${cellPattern}(?![A-Za-z0-9_(])`, 'g'));
  }
  const literalFree = formula.replace(/"[^"]*"/g, '""');
  for (const pattern of patterns) {
    for (const match of literalFree.matchAll(pattern)) {
      const startCol = columnNumber(match[1]);
      const startRow = Number(match[2]);
      const endCol = match[3] ? columnNumber(match[3]) : startCol;
      const endRow = match[4] ? Number(match[4]) : startRow;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          addresses.push(`${columnLetter(col)}${row}`);
        }
      }
    }
  }
  return new Set(addresses);
}

function splitAddress(address) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  return { col: columnNumber(match[1]), row: Number(match[2]) };
}

function columnNumber(letters) {
  return [...letters].reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function columnLetter(number) {
  let letters = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    number = Math.floor((number - 1) / 26);
  }
  return letters;
}

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function truncate(text) {
  return text.length > 120 ? `${text.slice(0, 120)}…` : text;
}
