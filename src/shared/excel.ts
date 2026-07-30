import ExcelJS from 'exceljs';
import type { PluginConfig, ReportContext, SourceRows } from './types';

export async function createInitialTemplate(config: PluginConfig, reportName: string): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'kintone Excel report plugin';
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  const settings = workbook.addWorksheet('設定');
  settings.columns = [
    { header: '項目', key: 'label', width: 22 },
    { header: '値', key: 'value', width: 34 }
  ];
  settings.addRows([
    { label: '帳票名', value: reportName },
    { label: '対象店舗', value: '' },
    { label: '基準日', value: '' },
    { label: '対象期間開始', value: '' },
    { label: '対象期間終了', value: '' },
    { label: '出力日', value: '' },
    { label: '出力者', value: '' }
  ]);
  settings.getRow(1).font = { bold: true };

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
  }

  const summary = workbook.addWorksheet('集計');
  summary.getCell('A1').value = 'ここにExcel側で数式・書式・レイアウトを作成してください';
  summary.getCell('A1').font = { color: { argb: 'FF667085' } };

  return workbook.xlsx.writeBuffer();
}

export async function fillReportTemplate(templateBuffer: ArrayBuffer, context: ReportContext, sourceRows: SourceRows[]): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.modified = new Date();

  writeSettingsSheet(workbook, context);

  for (const sourceResult of sourceRows) {
    writeSourceSheet(workbook, sourceResult);
  }

  return workbook.xlsx.writeBuffer();
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

function writeSettingsSheet(workbook: ExcelJS.Workbook, context: ReportContext): void {
  const worksheet = workbook.getWorksheet('設定') ?? workbook.addWorksheet('設定');
  const values = [
    ['帳票ID', context.reportId],
    ['帳票名', context.reportName],
    ['対象店舗', context.store],
    ['基準日', context.baseDate],
    ['対象期間開始', context.periodStart],
    ['対象期間終了', context.periodEnd],
    ['出力日', context.exportedAt],
    ['出力者', context.exporter]
  ];

  values.forEach(([label, value], index) => {
    const rowNumber = index + 1;
    worksheet.getCell(rowNumber, 1).value = label;
    worksheet.getCell(rowNumber, 2).value = value;
  });
}

function writeSourceSheet(workbook: ExcelJS.Workbook, sourceResult: SourceRows): void {
  const { source, rows } = sourceResult;
  const worksheet = workbook.getWorksheet(source.sheetName) ?? workbook.addWorksheet(source.sheetName);
  const headerRow = worksheet.getRow(1);

  source.fields.forEach((field, index) => {
    headerRow.getCell(index + 1).value = field.label || field.code;
  });
  applyColumnFormats(worksheet, source);
  headerRow.font = { bold: true };

  const existingRows = worksheet.rowCount;
  if (existingRows > 1) {
    worksheet.spliceRows(2, existingRows - 1);
  }

  rows.forEach((row, rowIndex) => {
    const worksheetRow = worksheet.getRow(rowIndex + 2);
    source.fields.forEach((field, columnIndex) => {
      worksheetRow.getCell(columnIndex + 1).value = normalizeCellValue(row[field.code]);
    });
    worksheetRow.commit();
  });
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

function sanitizeFileName(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, '_').trim();
}
