import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import type { PluginConfig, ReportContext, SourceAppConfig, SourceRows } from './types';

const settingsSheetName = '設定';
const summarySheetName = '集計';
const formulaPlaceholderPrefix = '__PL_FORMULA__:';
const settingsRows = [
  { label: '帳票ID', value: (context: ReportContext, reportName: string) => context.reportId || '' },
  { label: '帳票名', value: (context: ReportContext, reportName: string) => context.reportName || reportName },
  { label: '対象店舗', value: (context: ReportContext) => context.store },
  { label: '基準日', value: (context: ReportContext) => context.baseDate },
  { label: '対象期間開始', value: (context: ReportContext) => context.periodStart },
  { label: '対象期間終了', value: (context: ReportContext) => context.periodEnd },
  { label: '出力日', value: (context: ReportContext) => context.exportedAt },
  { label: '出力者', value: (context: ReportContext) => context.exporter }
];

export async function createInitialTemplate(config: PluginConfig, reportName: string): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'kintone Excel report plugin';
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  const settings = workbook.addWorksheet(settingsSheetName);
  settings.getColumn(1).width = 22;
  settings.getColumn(2).width = 34;
  settingsRows.forEach((row, index) => {
    const rowNumber = index + 1;
    settings.getCell(rowNumber, 1).value = row.label;
    settings.getCell(rowNumber, 2).value = row.value(emptyReportContext(), reportName);
  });

  const usedTableNames = new Set<string>();
  for (const source of config.sources) {
    const worksheet = workbook.addWorksheet(source.sheetName);
    worksheet.columns = source.fields.map((field) => ({
      header: field.label || field.code,
      key: field.code,
      width: Math.max(14, Math.min(32, (field.label || field.code).length + 6))
    }));
    applyColumnFormats(worksheet, source);
    worksheet.getRow(1).font = { bold: true };
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    applySourceTable(worksheet, { source, rows: [], periodStart: '', periodEnd: '', query: '' }, usedTableNames);
  }

  const summary = workbook.addWorksheet(summarySheetName);
  summary.getCell('A1').value = 'ここにExcel側で数式・書式・レイアウトを作成してください';
  summary.getCell('A1').font = { color: { argb: 'FF667085' } };

  return ensureAutomaticRecalculation(await workbook.xlsx.writeBuffer());
}

export async function fillReportTemplate(templateBuffer: ArrayBuffer, context: ReportContext, sourceRows: SourceRows[]): Promise<ArrayBuffer> {
  const workbook = await loadTemplateWorkbook(templateBuffer);
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.modified = new Date();

  validateTemplateWorkbook(workbook, sourceRows);
  writeSettingsSheet(workbook, context, sourceRows);

  const usedTableNames = new Set<string>();
  for (const sourceResult of sourceRows) {
    writeSourceSheet(workbook, sourceResult, usedTableNames);
  }
  materializeTemplateFormulas(workbook);
  hideTechnicalSheets(workbook, sourceRows);

  return ensureAutomaticRecalculation(await workbook.xlsx.writeBuffer());
}

export async function validateReportTemplate(templateBuffer: ArrayBuffer, sources: SourceAppConfig[]): Promise<void> {
  const workbook = await loadTemplateWorkbook(templateBuffer);

  validateTemplateWorkbook(
    workbook,
    sources.map((source) => ({
      source,
      rows: [],
      periodStart: '',
      periodEnd: '',
      query: ''
    }))
  );
}

async function loadTemplateWorkbook(templateBuffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(templateBuffer);
  } catch {
    throw new Error('完成版テンプレートExcelを読み込めません。xlsx形式か確認してください。');
  }
  return workbook;
}

export function downloadWorkbook(buffer: ArrayBuffer, fileName: string): void {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

export function buildFileName(context: ReportContext): string {
  const parts = [context.reportName || context.reportId || '帳票', context.store || '全店舗', context.baseDate]
    .filter(Boolean)
    .map((part) => sanitizeFileName(part));
  return `${parts.join('_')}.xlsx`;
}

function writeSettingsSheet(workbook: ExcelJS.Workbook, context: ReportContext, sourceRows: SourceRows[]): void {
  const worksheet = workbook.getWorksheet(settingsSheetName);
  if (!worksheet) {
    throw new Error(`テンプレートに「${settingsSheetName}」シートがありません。`);
  }

  const existingRows = worksheet.rowCount;

  settingsRows.forEach((row, index) => {
    const rowNumber = index + 1;
    worksheet.getCell(rowNumber, 1).value = row.label;
    worksheet.getCell(rowNumber, 2).value = row.value(context, context.reportName);
  });

  const detailStartRow = settingsRows.length + 2;
  const detailRows: Array<[string, string | number]> = [
    ['取得元アプリ数', sourceRows.length]
  ];
  sourceRows.forEach((sourceResult, index) => {
    const prefix = `取得元${index + 1}`;
    detailRows.push(
      [`${prefix}名`, sourceResult.source.label],
      [`${prefix}アプリID`, sourceResult.source.appId],
      [`${prefix}シート名`, sourceResult.source.sheetName],
      [`${prefix}テーブル名`, sourceResult.source.tableName || sourceResult.source.key],
      [`${prefix}対象期間開始`, sourceResult.periodStart],
      [`${prefix}対象期間終了`, sourceResult.periodEnd],
      [`${prefix}取得件数`, sourceResult.rows.length],
      [`${prefix}クエリ`, sourceResult.query]
    );
  });

  detailRows.forEach(([label, value], index) => {
    const rowNumber = detailStartRow + index;
    worksheet.getCell(rowNumber, 1).value = label;
    worksheet.getCell(rowNumber, 2).value = value;
  });

  clearLeftoverRows(worksheet, detailStartRow + detailRows.length, existingRows);
}

function writeSourceSheet(workbook: ExcelJS.Workbook, sourceResult: SourceRows, usedTableNames: Set<string>): void {
  const { source, rows } = sourceResult;
  const worksheet = workbook.getWorksheet(source.sheetName);
  if (!worksheet) {
    throw new Error(`テンプレートに元データ用シート「${source.sheetName}」がありません。`);
  }
  const headerRow = worksheet.getRow(1);

  source.fields.forEach((field, index) => {
    headerRow.getCell(index + 1).value = field.label || field.code;
  });
  applyColumnFormats(worksheet, source);
  headerRow.font = { bold: true };

  const existingRows = worksheet.rowCount;

  clearLeftoverRows(worksheet, rows.length + 2, existingRows);

  applySourceTable(worksheet, sourceResult, usedTableNames);
}

function clearLeftoverRows(worksheet: ExcelJS.Worksheet, fromRow: number, existingLastRow: number): void {
  for (let rowNumber = fromRow; rowNumber <= existingLastRow; rowNumber++) {
    worksheet.getRow(rowNumber).values = [];
  }
}

function materializeTemplateFormulas(workbook: ExcelJS.Workbook): void {
  workbook.eachSheet((worksheet) => {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value === 'string' && cell.value.startsWith(formulaPlaceholderPrefix)) {
          cell.value = { formula: cell.value.slice(formulaPlaceholderPrefix.length).replace(/^=/, '') };
        } else if (typeof cell.result === 'string' && cell.result.startsWith(formulaPlaceholderPrefix) && cell.formula) {
          cell.value = { formula: cell.formula };
        }
      });
    });
  });
}

async function ensureAutomaticRecalculation(workbookBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(workbookBuffer);
  const workbookXml = zip.file('xl/workbook.xml');
  if (!workbookXml) {
    throw new Error('Excelブック定義を読み込めません。');
  }

  const calcProperties = '<calcPr calcId="171027" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
  const xml = await workbookXml.async('string');
  const rewritten = /<calcPr\b[^>]*\/>/.test(xml)
    ? xml.replace(/<calcPr\b[^>]*\/>/, calcProperties)
    : xml.replace('</workbook>', `${calcProperties}</workbook>`);

  zip.file('xl/workbook.xml', rewritten);
  return zip.generateAsync({ type: 'arraybuffer' });
}

function hideTechnicalSheets(workbook: ExcelJS.Workbook, sourceRows: SourceRows[]): void {
  const names = new Set([settingsSheetName, summarySheetName, ...sourceRows.map(({ source }) => source.sheetName)]);
  const reportSheet = workbook.worksheets.find((worksheet) => !names.has(worksheet.name));
  if (!reportSheet) {
    return;
  }

  reportSheet.state = 'visible';
  const activeTab = workbook.worksheets.indexOf(reportSheet);
  workbook.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: activeTab, activeTab, visibility: 'visible' }];
  names.forEach((name) => {
    const worksheet = workbook.getWorksheet(name);
    if (worksheet) worksheet.state = 'hidden';
  });
}

function validateTemplateWorkbook(workbook: ExcelJS.Workbook, sourceRows: SourceRows[]): void {
  const errors: string[] = [];

  if (!workbook.getWorksheet(settingsSheetName)) {
    errors.push(`テンプレートに「${settingsSheetName}」シートがありません。`);
  }
  if (!workbook.getWorksheet(summarySheetName)) {
    errors.push(`テンプレートに「${summarySheetName}」シートがありません。`);
  }

  sourceRows.forEach(({ source }) => {
    const worksheet = workbook.getWorksheet(source.sheetName);
    if (!worksheet) {
      errors.push(`テンプレートに元データ用シート「${source.sheetName}」がありません。`);
      return;
    }

    const headerValues = worksheet.getRow(1).values;
    const headerLabels = new Set(
      (Array.isArray(headerValues) ? headerValues : [])
        .slice(1)
        .map((value: unknown) => String(value ?? '').trim())
        .filter(Boolean)
    );
    source.fields.forEach((field) => {
      const label = field.label || field.code;
      if (!headerLabels.has(label)) {
        errors.push(`元データ用シート「${source.sheetName}」に列「${label}」がありません。`);
      }
    });
  });

  if (errors.length) {
    throw new Error(['テンプレートの検証でエラーが見つかりました。', ...errors].join('\n'));
  }
}

function emptyReportContext(): ReportContext {
  return {
    reportId: '',
    reportName: '',
    store: '',
    baseDate: '',
    periodStart: '',
    periodEnd: '',
    exportedAt: '',
    exporter: ''
  };
}

function normalizeCellValue(value: unknown): ExcelJS.CellValue {
  if (value == null) {
    return '';
  }

  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string') {
    return value;
  }

  if (value instanceof Date) {
    return value;
  }

  return String(value);
}

function applyColumnFormats(worksheet: ExcelJS.Worksheet, source: SourceRows['source']): void {
  source.fields.forEach((field, index) => {
    const column = worksheet.getColumn(index + 1);

    if (field.type === 'number') {
      column.numFmt = '#,##0.########';
    }

    if (field.type === 'date') {
      column.numFmt = 'yyyy-mm-dd';
    }

    if (field.type === 'datetime') {
      column.numFmt = 'yyyy-mm-dd hh:mm';
    }
  });
}

function applySourceTable(worksheet: ExcelJS.Worksheet, sourceResult: SourceRows, usedTableNames: Set<string>): void {
  const { source, rows } = sourceResult;
  if (!source.fields.length) {
    return;
  }

  const tableName = uniqueTableName(source.tableName || source.key || source.sheetName, usedTableNames);
  removeTableAtSourceRange(worksheet, tableName);

  worksheet.addTable({
    name: tableName,
    displayName: tableName,
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: {
      theme: 'TableStyleLight9',
      showRowStripes: false
    },
    columns: source.fields.map((field) => ({
      name: field.label || field.code,
      filterButton: false
    })),
    rows: rows.length
      ? rows.map((row) => source.fields.map((field) => normalizeCellValue(row[field.code])))
      : [source.fields.map(() => '')]
  });
}

function removeTableAtSourceRange(worksheet: ExcelJS.Worksheet, tableName: string): void {
  const getTables = (worksheet as unknown as { getTables?: () => ExcelJS.Table[] }).getTables;
  const tables = getTables?.call(worksheet) ?? [];
  tables
    .filter((table) => table.name === tableName || table.ref === 'A1')
    .forEach((table) => worksheet.removeTable(table.name));
}

function uniqueTableName(value: string, usedTableNames: Set<string>): string {
  const baseName = sanitizeExcelTableName(value);
  let name = baseName;
  let sequence = 2;
  while (usedTableNames.has(name.toLowerCase())) {
    name = `${baseName}_${sequence}`;
    sequence += 1;
  }
  usedTableNames.add(name.toLowerCase());
  return name;
}

function sanitizeExcelTableName(value: string): string {
  // Excelのテーブル名(定義された名前)は日本語などのUnicode文字を許容するため、
  // 半角英数字だけに絞らず \p{L}(文字)・\p{N}(数字)・_ を残す。
  const sanitized = value
    .trim()
    .replace(/[^\p{L}\p{N}_]+/gu, '_')
    .replace(/^_+|_+$/g, '');
  const candidate = sanitized || 'source';
  const looksLikeCellReference = /^[A-Za-z]{1,3}[0-9]{1,7}$/.test(candidate);
  const withPrefix = /^[\p{L}_]/u.test(candidate) && !looksLikeCellReference ? candidate : `tbl_${candidate}`;
  return withPrefix.slice(0, 255);
}

function sanitizeFileName(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, '_').trim();
}
