// 各プロジェクトの実テンプレートを、実際の出力処理（fillReportTemplate）に通して検査する回帰テストの共通部分。
// scripts/pl-template-test.mjs（PL Management）と scripts/sales-template-test.mjs（営業分析）が使う。
//
// - テンプレートと取得元アプリ設定JSONは各プロジェクトのフォルダから読み込み、このリポジトリには
//   コミットしない（業務情報を含むため）。フォルダが無い環境ではスキップする。
// - 取得データは実データではなく、各ランナーが渡す rowsFor で作る合成データを使う。
// - useExcel が真なら、出力ファイルを実際のExcel（COM）で開いて全再計算し、ファイル破損や
//   #REF!・#NAME?などのエラーセルが無いこと、帳票シートのグラフ数も確認する（Windows＋Excelが必要）。
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

const formulaPlaceholderPrefix = '__PL_FORMULA__:';
const settingsSheetName = '設定';
const summarySheetName = '集計';
const drawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';

// options:
//   projectName: 表示用のプロジェクト名（例: 'PL Management'）
//   projectRoot: プロジェクトのフォルダ。無ければスキップする
//   cases: [{ name, template, sources, baseDate, expectedPeriod, store?, charts? }]（template・sourcesは projectRoot/outputs からの相対パス）
//   rowsFor(source, range, testCase, helpers): 取得元ごとの合成データ（fillReportTemplateに渡す行）を返す
//   reportId, exportedAt: 設定シートへ書く値
//   useExcel: 実際のExcelでも確認する
export async function runTemplateTests({ projectName, projectRoot, cases, rowsFor, reportId, exportedAt, useExcel }) {
  if (!existsSync(projectRoot)) {
    console.log(`${projectName}のフォルダが見つからないため、実テンプレートテストをスキップしました: ${projectRoot}`);
    return;
  }

  const stamp = `${process.pid}-${Date.now()}`;
  const queryOutput = path.join(tmpdir(), `krp-real-query-${stamp}.mjs`);
  const excelOutput = path.join(tmpdir(), `krp-real-excel-${stamp}.cjs`);
  const configOutput = path.join(tmpdir(), `krp-real-config-${stamp}.mjs`);
  const failures = [];

  try {
    await Promise.all([
      build({ entryPoints: ['src/shared/kintoneApi.ts'], bundle: true, platform: 'node', format: 'esm', outfile: queryOutput }),
      build({ entryPoints: ['src/shared/excel.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: excelOutput }),
      build({ entryPoints: ['src/shared/config.ts'], bundle: true, platform: 'node', format: 'esm', outfile: configOutput })
    ]);

    const { buildSourceQuery, sourceDateRange, mergeSourceRanges, fieldDisplayValue } = await import(pathToFileURL(queryOutput).href);
    const { fillReportTemplate, validateReportTemplate } = require(excelOutput);
    const { parseTemplateSources } = await import(pathToFileURL(configOutput).href);

    const outputDir = useExcel ? await mkdtemp(path.join(tmpdir(), 'krp-real-output-')) : '';
    const excelTargets = [];

    for (const testCase of cases) {
      let caseFailures = 0;
      const fail = (message) => {
        caseFailures += 1;
        if (caseFailures <= 20) failures.push(`[${testCase.name}] ${message}`);
        else if (caseFailures === 21) failures.push(`[${testCase.name}] ……以降の失敗は省略`);
      };
      const templatePath = path.join(projectRoot, 'outputs', testCase.template);
      const sourcesPath = path.join(projectRoot, 'outputs', testCase.sources);
      if (!existsSync(templatePath) || !existsSync(sourcesPath)) {
        fail(`テンプレートまたは取得元アプリ設定JSONが見つかりません。${projectName}側で差し替えた場合は、このテストの一覧を更新してください。\n  ${templatePath}\n  ${sourcesPath}`);
        continue;
      }

      // Kintoneに保存された取得元アプリ設定JSON（文字列）をパースする、実際の入力経路を通す。
      const sources = parseTemplateSources(await readFile(sourcesPath, 'utf8'), []);
      const templateBuffer = toArrayBuffer(await readFile(templatePath));
      await validateReportTemplate(templateBuffer, sources);

      const store = testCase.store ?? '';
      const sourceRows = sources.map((source) => {
        const range = sourceDateRange(source, testCase.baseDate);
        return {
          source,
          rows: rowsFor(source, range, testCase, { fieldDisplayValue }),
          periodStart: range.start,
          periodEnd: range.end,
          query: buildSourceQuery(source, store, testCase.baseDate)
        };
      });
      const wholeRange = mergeSourceRanges(sourceRows, testCase.baseDate);
      const context = {
        reportId,
        reportName: testCase.name,
        store,
        baseDate: testCase.baseDate,
        periodStart: wholeRange.periodStart,
        periodEnd: wholeRange.periodEnd,
        exportedAt,
        exporter: 'テスト'
      };
      if (context.periodStart !== testCase.expectedPeriod[0] || context.periodEnd !== testCase.expectedPeriod[1]) {
        fail(`対象期間が想定と違います。expected: ${testCase.expectedPeriod.join('〜')} actual: ${context.periodStart}〜${context.periodEnd}`);
      }

      const outputBuffer = await fillReportTemplate(templateBuffer, context, sourceRows);

      const template = new ExcelJS.Workbook();
      await template.xlsx.load(await stripChartDrawings(templateBuffer));
      const output = new ExcelJS.Workbook();
      await output.xlsx.load(await stripChartDrawings(outputBuffer));

      checkFormulas(template, output, fail);
      checkSettingsSheet(template, output, context, fail);
      checkSheetVisibility(template, output, sourceRows, fail);
      checkSourceTables(output, sourceRows, fail);
      await checkWorkbookXml(outputBuffer, fail);
      const chartCount = await checkTemplateCharts(templateBuffer, outputBuffer, fail);
      if (chartCount !== (testCase.charts ?? 0)) {
        fail(`テンプレートのグラフ数が想定と違います。expected: ${testCase.charts ?? 0} actual: ${chartCount}`);
      }

      if (useExcel) {
        const outputPath = path.join(outputDir, `case${excelTargets.length + 1}.xlsx`);
        await writeFile(outputPath, Buffer.from(outputBuffer));
        excelTargets.push({ name: testCase.name, path: outputPath, charts: testCase.charts ?? 0 });
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
      throw new Error([`${projectName}実テンプレートテストに失敗しました。`, ...failures].join('\n'));
    }
    console.log(`${projectName} template tests passed (${cases.length} templates${useExcel ? ', Excel verified' : ''}).`);
  } finally {
    await Promise.all([
      rm(queryOutput, { force: true }),
      rm(excelOutput, { force: true }),
      rm(configOutput, { force: true })
    ]);
  }
}

// ExcelJSは、openpyxlで置いたグラフ（接頭辞なしの描画）を含むファイルを読めない。検査用に読み込むときは、
// 描画への参照と描画の部品を外してから読み込む（グラフ自体は checkTemplateCharts でZIPのまま確かめる）。
// グラフのないファイルはそのまま返す。
async function stripChartDrawings(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const drawings = Object.keys(zip.files).filter((name) => /^xl\/drawings\/[^/]+\.xml$/.test(name));
  let changed = false;
  for (const name of drawings) {
    const xml = await zip.file(name).async('string');
    if (!/drawingml\/2006\/chart"/.test(xml)) continue;
    changed = true;
    zip.remove(name);
    zip.remove(relsOf(name));
  }
  if (!changed) return buffer;
  for (const sheetPath of (await sheetPaths(zip)).values()) {
    const rels = await relationships(zip, sheetPath);
    const removed = rels.filter((rel) => rel.type === drawingRelationshipType && !zip.file(resolve(sheetPath, rel.target)));
    if (!removed.length) continue;
    let sheetXml = await zip.file(sheetPath).async('string');
    let relsXml = await zip.file(relsOf(sheetPath)).async('string');
    for (const rel of removed) {
      sheetXml = sheetXml.replace(new RegExp(`<drawing [^>]*r:id="${rel.id}"[^>]*/>`), '');
      relsXml = relsXml.replace(rel.element, '');
    }
    zip.file(sheetPath, sheetXml);
    zip.file(relsOf(sheetPath), relsXml);
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}

// テンプレートの各シートのグラフが、出力ファイルの同名シートにそのまま残っていること（v0.17.0）。
// グラフのないテンプレートでは、出力ファイルにもグラフの部品が無いこと。返り値はテンプレートのグラフ数。
async function checkTemplateCharts(templateBuffer, outputBuffer, fail) {
  const template = await JSZip.loadAsync(templateBuffer);
  const output = await JSZip.loadAsync(outputBuffer);
  const outputSheets = await sheetPaths(output);
  let chartCount = 0;
  for (const [sheetName, sheetPath] of await sheetPaths(template)) {
    const templateCharts = await sheetCharts(template, sheetPath);
    if (!templateCharts.length) continue;
    chartCount += templateCharts.length;
    const outputCharts = outputSheets.has(sheetName) ? await sheetCharts(output, outputSheets.get(sheetName)) : [];
    if (outputCharts.length !== templateCharts.length) {
      fail(`「${sheetName}」シートのグラフが出力ファイルに残っていません。expected: ${templateCharts.length} actual: ${outputCharts.length}`);
      continue;
    }
    templateCharts.forEach((chart, index) => {
      if (outputCharts[index].xml !== chart.xml) fail(`「${sheetName}」シートの${index + 1}つ目のグラフの内容がテンプレートと違います。`);
    });
  }
  if (!chartCount && Object.keys(output.files).some((name) => /^xl\/charts\//.test(name))) {
    fail('グラフのないテンプレートの出力ファイルに、グラフの部品が入っています。');
  }
  return chartCount;
}

async function sheetCharts(zip, sheetPath) {
  const charts = [];
  for (const rel of await relationships(zip, sheetPath)) {
    if (rel.type !== drawingRelationshipType) continue;
    const drawingPath = resolve(sheetPath, rel.target);
    const drawingXml = (await zip.file(drawingPath)?.async('string')) ?? '';
    const chartRels = new Map((await relationships(zip, drawingPath)).map((chartRel) => [chartRel.id, chartRel]));
    // 描画の中の並び順（アンカー順）でグラフを並べる。
    for (const match of drawingXml.matchAll(/<(?:[\w]+:)?chart\b[^>]*\br:id="([^"]+)"/g)) {
      const chartRel = chartRels.get(match[1]);
      if (chartRel?.type !== chartRelationshipType) continue;
      charts.push({ xml: (await zip.file(resolve(drawingPath, chartRel.target))?.async('string')) ?? null });
    }
  }
  return charts;
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

// 出力ファイルを実際のExcelで開き、全再計算して、表示中の帳票シートのエラーセルとグラフを数える。
// Excelが修復を必要とするファイル（v0.16.9のような破損）は、自動化では開けずに失敗する。
async function checkWithExcel(targets, outputDir, failures) {
  const scriptPath = path.join(outputDir, 'check.ps1');
  const listPath = path.join(outputDir, 'targets.json');
  await writeFile(listPath, JSON.stringify(targets), 'utf8');
  await writeFile(
    scriptPath,
    `$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$targets = Get-Content -LiteralPath '${listPath}' -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$results = @()
try {
  foreach ($target in $targets) {
    $result = [ordered]@{ name = $target.name; opened = $false; formulaCount = 0; charts = 0; errors = @{}; samples = @() }
    try {
      $workbook = $excel.Workbooks.Open($target.path, 0, $true)
      $result.opened = $true
      $excel.CalculateFullRebuild()
      foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Visible -ne -1) { continue }
        $result.charts += $sheet.ChartObjects().Count
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
  // PowerShellの出力の名前ではなく、順番で対応させる。
  targets.forEach((target, index) => {
    const result = results[index] ?? {};
    const errors = Object.entries(result.errors ?? {});
    const breakingErrors = errors.filter(([text]) => !dataDependentErrors.has(text));
    const note = errors.length ? `、データ依存のエラー ${JSON.stringify(Object.fromEntries(errors))}` : '';
    const chartNote = target.charts ? `、グラフ${result.charts}個` : '';
    if (!result.opened) {
      failures.push(`[${target.name}] Excelで出力ファイルを開けません（ファイル破損の可能性）: ${result.openError}`);
    } else if (!result.formulaCount) {
      failures.push(`[${target.name}] Excelで開いた帳票シートに数式が1件もありません（シート内容が破棄された可能性）。`);
    } else if (breakingErrors.length) {
      failures.push(`[${target.name}] Excelの全再計算後にエラーセルがあります: ${JSON.stringify(Object.fromEntries(breakingErrors))} 例: ${result.samples.join(', ')}`);
    } else if (result.charts !== target.charts) {
      failures.push(`[${target.name}] Excelで数えた帳票シートのグラフ数が想定と違います。expected: ${target.charts} actual: ${result.charts}`);
    } else {
      console.log(`  ${target.name}: Excelで開けて全再計算できた（数式${result.formulaCount}件${chartNote}${note}）`);
    }
  });
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
// 「イベント日設定!A2」のように、名前が「設定」で終わる別のシートへの参照は対象にしない。
function settingsReferences(formula, onSettingsSheet) {
  const addresses = [];
  const cellPattern = '\\$?([A-Z]{1,3})\\$?(\\d+)(?::\\$?([A-Z]{1,3})\\$?(\\d+))?';
  const patterns = [new RegExp(`(?<![\\p{L}\\p{N}_.])'?${settingsSheetName}'?!${cellPattern}`, 'gu')];
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

async function sheetPaths(zip) {
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const targets = new Map((await relationships(zip, 'xl/workbook.xml')).map((rel) => [rel.id, rel.target]));
  const sheets = new Map();
  for (const [element] of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = decodeXml(element.match(/\bname="([^"]*)"/)[1]);
    const id = element.match(/\br:id="([^"]*)"/)[1];
    sheets.set(name, resolve('xl/workbook.xml', targets.get(id)));
  }
  return sheets;
}

async function relationships(zip, partPath) {
  const xml = await zip.file(relsOf(partPath))?.async('string');
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map(([element]) => ({
    element,
    id: element.match(/\bId="([^"]*)"/)[1],
    type: element.match(/\bType="([^"]*)"/)[1],
    target: element.match(/\bTarget="([^"]*)"/)[1]
  }));
}

function relsOf(partPath) {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

function resolve(fromPart, target) {
  if (target.startsWith('/')) return target.slice(1);
  return path.posix.normalize(path.posix.join(path.posix.dirname(fromPart), target));
}

function decodeXml(value) {
  return value.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
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
