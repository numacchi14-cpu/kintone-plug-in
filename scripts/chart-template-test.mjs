// 完成版テンプレートに置いたグラフが出力ファイルに残ることを確かめるテスト（v0.17.0）。
//
// - テンプレートは scripts/fixtures/ の合成ファイル（業務データなし。作り方は make-chart-templates.mjs）。
//   - chart-excel.xlsx: Excelで置いた折れ線グラフ2つ（相対パスの関係、chartStyle・colorsつき）
//   - chart-openpyxl.xlsx: openpyxlで置いた折れ線グラフ1つ（描画・グラフへの関係が絶対パス）
//   - 上記に加え、別のシートに画像（ExcelJSがそのまま扱う描画）を足したテンプレートもテスト内で作る
// - 実際の出力処理（fillReportTemplate）に通し、出力ファイルの部品・関係・Content Types・<drawing>要素を確かめる。
// - --excel を付けると、出力ファイルを実際のExcel（COM）で開き、修復なしで開けること、グラフの数、
//   全再計算後のグラフの値が取得データどおりであることも確かめる（Windows＋Excelが必要）。
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import JSZip from 'jszip';

const require = createRequire(import.meta.url);
const useExcel = process.argv.includes('--excel');
const fixtureDir = path.join('scripts', 'fixtures');
const reportSheetName = '帳票';
const summarySheetName = '集計';
const contentTypes = {
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  chartStyle: 'application/vnd.ms-office.chartstyle+xml',
  chartColors: 'application/vnd.ms-office.chartcolorstyle+xml'
};
const relationshipTypes = {
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
};

const source = {
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
// 1〜12月の各月に2行。売上は月ごとに変える（グラフの値が取得データどおりか確かめるため）。
const rows = Array.from({ length: 12 }, (_, index) => [
  { date: `2026-${String(index + 1).padStart(2, '0')}-05`, store: 'テスト店', sales: 1000 * (index + 1) },
  { date: `2026-${String(index + 1).padStart(2, '0')}-20`, store: 'テスト店', sales: 500 }
]).flat();
const expectedSales = Array.from({ length: 12 }, (_, index) => 1000 * (index + 1) + 500);
const context = {
  reportId: 'chart_test',
  reportName: 'グラフテスト',
  store: 'テスト店',
  baseDate: '2026-12-31',
  periodStart: '2026-01-01',
  periodEnd: '2026-12-31',
  exportedAt: '2026-10-14T10:00:00',
  exporter: 'テスト'
};

const stamp = `${process.pid}-${Date.now()}`;
const excelOutput = path.join(tmpdir(), `krp-chart-excel-${stamp}.cjs`);
const chartsOutput = path.join(tmpdir(), `krp-chart-charts-${stamp}.mjs`);
const failures = [];

try {
  await Promise.all([
    build({ entryPoints: ['src/shared/excel.ts'], bundle: true, platform: 'node', format: 'cjs', outfile: excelOutput }),
    build({ entryPoints: ['src/shared/charts.ts'], bundle: true, platform: 'node', format: 'esm', outfile: chartsOutput })
  ]);
  const { fillReportTemplate, validateReportTemplate } = require(excelOutput);
  const { extractTemplateCharts } = await import(pathToFileURL(chartsOutput).href);

  const excelTemplate = await readFixture('chart-excel.xlsx');
  const cases = [
    { name: 'Excelで置いたグラフ', template: excelTemplate, charts: 2, chartStyles: true },
    { name: 'openpyxlで置いたグラフ（絶対パスの関係）', template: await readFixture('chart-openpyxl.xlsx'), charts: 1, chartStyles: false },
    { name: 'グラフ＋別シートの画像', template: await addPictureSheetDrawing(excelTemplate), charts: 2, chartStyles: true, picture: true }
  ];

  // グラフのないテンプレート（PL Managementのテンプレートすべて）は、ExcelJSへ渡す内容を一切変えない。
  const baseTemplate = await readFixture('chart-base.xlsx');
  const untouched = await extractTemplateCharts(baseTemplate);
  check(untouched.buffer === baseTemplate && untouched.charts.drawings.length === 0, 'グラフのないテンプレートは元のバッファのままExcelJSへ渡す');

  const outputDir = useExcel ? await mkdtemp(path.join(tmpdir(), 'krp-chart-output-')) : '';
  const excelTargets = [];
  for (const testCase of cases) {
    const fail = (message) => failures.push(`[${testCase.name}] ${message}`);
    await validateReportTemplate(testCase.template, [source]);
    const output = await fillReportTemplate(testCase.template, context, [
      { source, rows, periodStart: context.periodStart, periodEnd: context.periodEnd, query: '' }
    ]);
    await checkOutputCharts(testCase, output, fail);

    // 数式プレースホルダーも従来どおり変換されている。ExcelJSで読み直しての確認はしない（openpyxlの描画は
    // 接頭辞なしの名前空間で書かれており、ExcelJSはそのような描画を含むファイルを読めないため）。
    const outputZip = await JSZip.loadAsync(output);
    const summaryXml = await outputZip.file((await sheetPaths(outputZip)).get(summarySheetName)).async('string');
    if (!/<c r="B3"[^>]*><f>SUMIFS\(/.test(summaryXml)) {
      fail('集計!B3 の数式プレースホルダーが変換されていません。');
    }

    if (useExcel) {
      const outputPath = path.join(outputDir, `chart${excelTargets.length + 1}.xlsx`);
      await writeFile(outputPath, Buffer.from(output));
      excelTargets.push({ name: testCase.name, path: outputPath, charts: testCase.charts, picture: Boolean(testCase.picture) });
    }
    console.log(`  ${testCase.name}: 出力ファイルの部品・関係の検査が完了`);
  }

  // グラフと画像が同じ描画（同じシート）にあるテンプレートは、分かりやすいエラーにする。
  const mixed = await JSZip.loadAsync(excelTemplate);
  const drawingXml = await mixed.file('xl/drawings/drawing1.xml').async('string');
  mixed.file('xl/drawings/drawing1.xml', drawingXml.replace('</xdr:wsDr>', `${pictureAnchor('rId9')}</xdr:wsDr>`));
  let mixedError = '';
  try {
    await validateReportTemplate(await mixed.generateAsync({ type: 'arraybuffer' }), [source]);
  } catch (error) {
    mixedError = error.message;
  }
  check(mixedError.includes('グラフと画像') && mixedError.includes(reportSheetName), `グラフと画像が同じシートにあるテンプレートはエラー（実際: ${mixedError || 'エラーなし'}）`);

  if (useExcel && excelTargets.length) {
    await checkWithExcel(excelTargets, outputDir);
  }
  if (outputDir) {
    await rm(outputDir, { recursive: true, force: true });
  }
  if (failures.length) {
    throw new Error(['グラフ入りテンプレートのテストに失敗しました。', ...failures].join('\n'));
  }
  console.log(`Chart template tests passed (${cases.length} templates${useExcel ? ', Excel verified' : ''}).`);
} finally {
  await Promise.all([rm(excelOutput, { force: true }), rm(chartsOutput, { force: true })]);
}

// 出力ファイルの帳票シートに、テンプレートの描画・グラフがそのまま戻っていること。
async function checkOutputCharts(testCase, outputBuffer, fail) {
  const template = await JSZip.loadAsync(testCase.template);
  const output = await JSZip.loadAsync(outputBuffer);
  const outputTypes = await overrideTypes(output);

  const sheetPath = (await sheetPaths(output)).get(reportSheetName);
  const sheetXml = await output.file(sheetPath).async('string');
  const drawingElements = [...sheetXml.matchAll(/<drawing r:id="([^"]+)"\/>/g)];
  if (drawingElements.length !== 1) {
    fail(`帳票シートの<drawing>要素が1つではありません（${drawingElements.length}個）。`);
    return;
  }
  // スキーマ上の順序: pageMargins・pageSetup・headerFooter の後、legacyDrawing・tableParts・extLst の前。
  const drawingIndex = drawingElements[0].index;
  for (const before of ['sheetData', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'ignoredErrors']) {
    const index = sheetXml.search(new RegExp(`<${before}\\b`));
    if (index > drawingIndex) fail(`<drawing>が<${before}>より前にあります（スキーマ違反）。`);
  }
  for (const after of ['legacyDrawing', 'legacyDrawingHF', 'picture', 'tableParts', 'extLst']) {
    const index = sheetXml.search(new RegExp(`<${after}\\b`));
    if (index >= 0 && index < drawingIndex) fail(`<drawing>が<${after}>より後ろにあります（スキーマ違反）。`);
  }

  const sheetRelationships = await relationships(output, sheetPath);
  const drawingRelationship = sheetRelationships.find((relationship) => relationship.id === drawingElements[0][1]);
  if (drawingRelationship?.type !== relationshipTypes.drawing) {
    fail(`<drawing>の関係（${drawingElements[0][1]}）が帳票シートの関係にありません。`);
    return;
  }
  const drawingPath = resolve(sheetPath, drawingRelationship.target);
  const templateSheetPath = (await sheetPaths(template)).get(reportSheetName);
  const templateDrawing = (await relationships(template, templateSheetPath)).find((relationship) => relationship.type === relationshipTypes.drawing);
  const templateDrawingPath = resolve(templateSheetPath, templateDrawing.target);
  await expectSamePart(template, templateDrawingPath, output, drawingPath, contentTypes.drawing, outputTypes, fail);

  const templateCharts = new Map((await relationships(template, templateDrawingPath)).map((relationship) => [relationship.id, relationship]));
  const outputCharts = (await relationships(output, drawingPath)).filter((relationship) => relationship.type === relationshipTypes.chart);
  if (outputCharts.length !== testCase.charts) {
    fail(`描画から参照するグラフの数が想定と違います。expected: ${testCase.charts} actual: ${outputCharts.length}`);
  }
  for (const chart of outputCharts) {
    if (chart.target.startsWith('/')) fail(`描画からグラフへの関係が絶対パスのままです: ${chart.target}`);
    const templateChartPath = resolve(templateDrawingPath, templateCharts.get(chart.id).target);
    const chartPath = resolve(drawingPath, chart.target);
    await expectSamePart(template, templateChartPath, output, chartPath, contentTypes.chart, outputTypes, fail);

    const templateParts = new Map((await relationships(template, templateChartPath)).map((relationship) => [relationship.id, relationship]));
    const chartParts = await relationships(output, chartPath);
    if (testCase.chartStyles && chartParts.length !== 2) {
      fail(`${chartPath} のchartStyle・colorsの関係が2つではありません（${chartParts.length}個）。`);
    }
    for (const part of chartParts) {
      const expectedType = part.type.endsWith('/chartStyle') ? contentTypes.chartStyle : contentTypes.chartColors;
      await expectSamePart(
        template,
        resolve(templateChartPath, templateParts.get(part.id).target),
        output,
        resolve(chartPath, part.target),
        expectedType,
        outputTypes,
        fail
      );
    }
  }

  // ExcelJSが扱う画像（別シート）はそのまま残り、テンプレートのグラフの描画と部品名が重ならない。
  if (testCase.picture) {
    const summaryPath = (await sheetPaths(output)).get(summarySheetName);
    const pictureDrawing = (await relationships(output, summaryPath)).find((relationship) => relationship.type === relationshipTypes.drawing);
    if (!pictureDrawing) {
      fail('別シートの画像の描画が出力ファイルにありません。');
    } else if (resolve(summaryPath, pictureDrawing.target) === drawingPath) {
      fail('画像の描画とグラフの描画が同じ部品を指しています。');
    }
  }
}

async function expectSamePart(template, templatePath, output, outputPath, expectedType, outputTypes, fail) {
  const expected = await template.file(templatePath)?.async('string');
  const actual = await output.file(outputPath)?.async('string');
  if (actual == null) {
    fail(`出力ファイルに ${outputPath}（テンプレートの ${templatePath}）がありません。`);
    return;
  }
  if (actual !== expected) fail(`${outputPath} の内容がテンプレートの ${templatePath} と違います。`);
  if (outputTypes.get(outputPath) !== expectedType) {
    fail(`${outputPath} のContent Typesが想定と違います。expected: ${expectedType} actual: ${outputTypes.get(outputPath)}`);
  }
}

// テンプレートの集計シートに画像を1つ置く（ExcelJSが読み書きする、グラフを含まない描画）。
async function addPictureSheetDrawing(templateBuffer) {
  const zip = await JSZip.loadAsync(templateBuffer);
  const summaryPath = (await sheetPaths(zip)).get(summarySheetName);
  const relsPath = relsOf(summaryPath);
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  );
  zip.file('xl/media/image1.png', png);
  zip.file(
    'xl/drawings/drawing9.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${pictureAnchor('rId1')}</xdr:wsDr>`
  );
  zip.file(
    'xl/drawings/_rels/drawing9.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationshipTypes.image}" Target="../media/image1.png"/></Relationships>`
  );
  const existing = (await zip.file(relsPath)?.async('string')) ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  zip.file(relsPath, existing.replace('</Relationships>', `<Relationship Id="rIdPicture" Type="${relationshipTypes.drawing}" Target="../drawings/drawing9.xml"/></Relationships>`));
  const sheetXml = await zip.file(summaryPath).async('string');
  const sheetWithNamespace = sheetXml.includes('xmlns:r=')
    ? sheetXml
    : sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  const insertAt = sheetWithNamespace.search(/<(legacyDrawing|tableParts|extLst)\b|<\/worksheet>/);
  zip.file(summaryPath, `${sheetWithNamespace.slice(0, insertAt)}<drawing r:id="rIdPicture"/>${sheetWithNamespace.slice(insertAt)}`);
  const types = await zip.file('[Content_Types].xml').async('string');
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      `${types.includes('Extension="png"') ? '' : '<Default Extension="png" ContentType="image/png"/>'}<Override PartName="/xl/drawings/drawing9.xml" ContentType="${contentTypes.drawing}"/></Types>`
    )
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

function pictureAnchor(relationshipId) {
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="190500" cy="190500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="90" name="図 90"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
}

// 出力ファイルを実際のExcelで開き、帳票シートのグラフの数と、全再計算後のグラフの値を確かめる。
async function checkWithExcel(targets, outputDir) {
  const scriptPath = path.join(outputDir, 'check-charts.ps1');
  const listPath = path.join(outputDir, 'targets.json');
  await writeFile(listPath, JSON.stringify(targets), 'utf8');
  // シート名（日本語）はJSONで渡し、スクリプト本体はASCIIだけにする（Windows PowerShell 5.1の文字コード対策）。
  await writeFile(
    scriptPath,
    `$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$targets = Get-Content -LiteralPath '${listPath}' -Raw -Encoding UTF8 | ConvertFrom-Json
$names = Get-Content -LiteralPath '${listPath}.names' -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$results = @()
try {
  foreach ($target in $targets) {
    $result = [ordered]@{ name = $target.name; opened = $false; charts = 0; series = @(); pictures = 0 }
    try {
      $workbook = $excel.Workbooks.Open($target.path, 0, $true)
      $result.opened = $true
      $excel.CalculateFullRebuild()
      $sheet = $workbook.Worksheets.Item($names.report)
      $result.charts = $sheet.ChartObjects().Count
      foreach ($chartObject in $sheet.ChartObjects()) {
        $series = $chartObject.Chart.SeriesCollection(1)
        $result.series += [pscustomobject]@{ formula = [string]$series.Formula; values = @($series.Values) }
      }
      $summary = $workbook.Worksheets.Item($names.summary)
      foreach ($shape in $summary.Shapes) { if ($shape.Type -eq 13) { $result.pictures += 1 } }
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
$results | ConvertTo-Json -Depth 6 -Compress
`,
    'utf8'
  );
  await writeFile(`${listPath}.names`, JSON.stringify({ report: reportSheetName, summary: summarySheetName }), 'utf8');

  console.log('  Excelで出力ファイルを開いてグラフを確かめています...');
  const stdout = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath], {
    encoding: 'utf8',
    maxBuffer: 16 * 1024 * 1024
  });
  const results = [JSON.parse(stdout.trim())].flat();
  // PowerShellの標準出力は日本語が化けることがあるため、名前ではなく順番で対応させる。
  for (const [index, target] of targets.entries()) {
    const result = results[index];
    const fail = (message) => failures.push(`[${target.name}] ${message}`);
    if (!result?.opened) {
      fail(`Excelで出力ファイルを開けません（ファイル破損の可能性）: ${result?.openError}`);
      continue;
    }
    if (result.charts !== target.charts) {
      fail(`Excelで数えた帳票シートのグラフ数が想定と違います。expected: ${target.charts} actual: ${result.charts}`);
    }
    const salesSeries = result.series.find((series) => series.formula.includes(`${summarySheetName}!$B$3:$M$3`));
    if (!salesSeries) {
      fail(`集計シートの売上の範囲を参照するグラフがありません: ${result.series.map((series) => series.formula).join(' / ')}`);
    } else if (JSON.stringify(salesSeries.values.map(Number)) !== JSON.stringify(expectedSales)) {
      fail(`全再計算後のグラフの値が取得データと違います。expected: ${expectedSales.join(',')} actual: ${salesSeries.values.join(',')}`);
    }
    if (target.picture && result.pictures !== 1) {
      fail(`別シートの画像がExcelで見つかりません（${result.pictures}個）。`);
    }
    if (!failures.some((message) => message.startsWith(`[${target.name}]`))) {
      console.log(`  ${target.name}: Excelで開けてグラフ${result.charts}個、全再計算後の値も一致`);
    }
  }
}

async function sheetPaths(zip) {
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const targets = new Map((await relationships(zip, 'xl/workbook.xml')).map((relationship) => [relationship.id, relationship.target]));
  const sheets = new Map();
  for (const [element] of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = element.match(/\bname="([^"]*)"/)[1];
    const id = element.match(/\br:id="([^"]*)"/)[1];
    sheets.set(name, resolve('xl/workbook.xml', targets.get(id)));
  }
  return sheets;
}

async function relationships(zip, partPath) {
  const xml = await zip.file(relsOf(partPath))?.async('string');
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map(([element]) => ({
    id: element.match(/\bId="([^"]*)"/)[1],
    type: element.match(/\bType="([^"]*)"/)[1],
    target: element.match(/\bTarget="([^"]*)"/)[1]
  }));
}

async function overrideTypes(zip) {
  const xml = await zip.file('[Content_Types].xml').async('string');
  return new Map(
    [...xml.matchAll(/<Override\b[^>]*>/g)].map(([element]) => [
      element.match(/\bPartName="\/([^"]*)"/)[1],
      element.match(/\bContentType="([^"]*)"/)[1]
    ])
  );
}

function relsOf(partPath) {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

function resolve(fromPart, target) {
  if (target.startsWith('/')) return target.slice(1);
  return path.posix.normalize(path.posix.join(path.posix.dirname(fromPart), target));
}

async function readFixture(name) {
  const buffer = await readFile(path.join(fixtureDir, name));
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function check(condition, label) {
  if (!condition) failures.push(label);
}
