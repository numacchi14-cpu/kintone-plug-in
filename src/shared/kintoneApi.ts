import type {
  KintoneRecord,
  PluginConfig,
  SourceAppConfig,
  SourceFieldConfig,
  SourceFieldValueType
} from './types';
import { calculateDateRange } from './dateRules';

declare const kintone: any;

interface KintoneFormField {
  code?: string;
  label?: string;
  type?: string;
  fields?: Record<string, KintoneFormField>;
}

export interface SourceFieldImportResult {
  fields: SourceFieldConfig[];
  skippedCount: number;
}

const unsupportedOutputFieldTypes = new Set([
  'FILE',
  'GROUP',
  'HR',
  'LABEL',
  'REFERENCE_TABLE',
  'SPACER',
  'SUBTABLE'
]);

export async function getSourceAppFields(appId: string): Promise<SourceFieldImportResult> {
  if (!/^\d+$/.test(appId)) {
    throw new Error('アプリIDは半角数字で入力してください。');
  }

  const response = await kintone.api(kintone.api.url('/k/v1/app/form/fields.json', true), 'GET', {
    app: appId,
    lang: 'ja'
  });
  const properties = (response.properties || {}) as Record<string, KintoneFormField>;
  const fields: SourceFieldConfig[] = [];
  let skippedCount = 0;

  Object.entries(properties).forEach(([propertyCode, property]) => {
    const code = property.code || propertyCode;
    const kintoneType = property.type || '';

    if (!code || unsupportedOutputFieldTypes.has(kintoneType)) {
      skippedCount += 1;
      return;
    }

    fields.push({
      code,
      label: property.label || code,
      type: sourceFieldValueType(kintoneType)
    });
  });

  return { fields, skippedCount };
}

export async function getAllRecords(
  appId: string,
  query: string,
  fields: string[],
  onProgress?: (count: number) => void
): Promise<KintoneRecord[]> {
  const records: KintoneRecord[] = [];
  const uniqueFields = Array.from(new Set(fields.filter(Boolean)));
  const cursorUrl = kintone.api.url('/k/v1/records/cursor.json', true);
  let cursorId = '';

  try {
    const cursor = await kintone.api(cursorUrl, 'POST', {
      app: appId,
      query,
      fields: uniqueFields,
      size: 500
    });

    cursorId = cursor.id;

    while (true) {
      const response = await kintone.api(cursorUrl, 'GET', { id: cursorId });
      records.push(...response.records);
      onProgress?.(records.length);

      if (!response.next) {
        cursorId = '';
        break;
      }
    }
  } catch (error) {
    if (cursorId) {
      await kintone.api(cursorUrl, 'DELETE', { id: cursorId }).catch(() => undefined);
    }
    throw error;
  }

  return records;
}

export async function findTemplateRecord(config: PluginConfig, reportId: string): Promise<KintoneRecord> {
  if (!config.templateAppId) {
    throw new Error('テンプレート管理アプリIDが未設定です。');
  }

  const query = `${fieldRef(config.templateReportIdField)} = "${escapeQueryValue(reportId)}" limit 1`;
  const response = await kintone.api(kintone.api.url('/k/v1/records.json', true), 'GET', {
    app: config.templateAppId,
    query
  });

  if (!response.records.length) {
    throw new Error(`帳票ID「${reportId}」のテンプレート設定が見つかりません。`);
  }

  return response.records[0];
}

export async function downloadKintoneFile(fileKey: string): Promise<ArrayBuffer> {
  const url = `${kintone.api.url('/k/v1/file.json', true)}?fileKey=${encodeURIComponent(fileKey)}`;

  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('GET', url);
    xhr.responseType = 'arraybuffer';
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        resolve(xhr.response);
      } else {
        reject(new Error(`Excelテンプレートの取得に失敗しました。HTTP ${xhr.status}`));
      }
    };
    xhr.onerror = () => reject(new Error('Excelテンプレートの取得中に通信エラーが発生しました。'));
    xhr.send();
  });
}

export async function uploadKintoneFile(file: File): Promise<string> {
  const url = kintone.api.url('/k/v1/file.json', true);
  const formData = new FormData();
  const requestToken = typeof kintone.getRequestToken === 'function' ? kintone.getRequestToken() : '';
  if (requestToken) {
    formData.append('__REQUEST_TOKEN__', requestToken);
  }
  formData.append('file', file);

  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', url);
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
    xhr.onload = () => {
      if (xhr.status < 200 || xhr.status >= 300) {
        reject(new Error(`完成版テンプレートのアップロードに失敗しました。HTTP ${xhr.status}`));
        return;
      }

      try {
        const response = JSON.parse(xhr.responseText || '{}') as { fileKey?: string };
        if (!response.fileKey) {
          reject(new Error('完成版テンプレートのアップロード結果にfileKeyがありません。'));
          return;
        }
        resolve(response.fileKey);
      } catch {
        reject(new Error('完成版テンプレートのアップロード結果を読み取れません。'));
      }
    };
    xhr.onerror = () => reject(new Error('完成版テンプレートのアップロード中に通信エラーが発生しました。'));
    xhr.send(formData);
  });
}

export function sourceDateRange(source: SourceAppConfig, baseDate: string): { start: string; end: string } {
  // 日付範囲フィルターを1つも持たないソース（例：配賦設定履歴のように全件取得するアプリ）は、
  // このソース自体には「対象期間」という概念がない。ここで基準日にフォールバックしてしまうと、
  // mergeSourceRangesが全ソースの期間開始・終了の最小・最大を取る際に、日付フィルター付きソース
  // の本当の期間ではなく、このダミーの基準日が採用されてしまう不具合があった（特に上半期・下半期
  // のように基準日が対象期間の外にある場合、期間終了が誤って基準日になってしまう）。
  // 空文字列を返し、呼び出し側のfilter(Boolean)で「対象期間なし」として除外させる。
  const ranges = source.filters
    .filter((filter) => filter.valueFrom === 'dateRange')
    .map((filter) => calculateDateRange(filter.dateRule || 'sameDay', baseDate));
  const starts = ranges.map((range) => range.start).sort();
  const ends = ranges.map((range) => range.end).sort();
  return {
    start: starts[0] || '',
    end: ends[ends.length - 1] || ''
  };
}

export function mergeSourceRanges(
  sourceRanges: Array<{ periodStart: string; periodEnd: string }>,
  fallbackDate: string
): { periodStart: string; periodEnd: string } {
  const starts = sourceRanges.map((source) => source.periodStart).filter(Boolean).sort();
  const ends = sourceRanges.map((source) => source.periodEnd).filter(Boolean).sort();
  return {
    periodStart: starts[0] || fallbackDate,
    periodEnd: ends[ends.length - 1] || fallbackDate
  };
}

export function buildSourceQuery(
  source: SourceAppConfig,
  storeValue: string,
  baseDate: string,
  outputRecord?: KintoneRecord
): string {
  const conditions: string[] = [];

  source.filters.forEach((filter) => {
    const condition = buildFilterCondition(filter, storeValue, baseDate, outputRecord);
    if (condition) {
      conditions.push(condition);
    }
  });

  const conditionQuery = conditions.length ? conditions.join(' and ') : '';
  const structuredSort = source.sorts
    .filter((sort) => sort.field)
    .map((sort) => `${fieldRef(sort.field)} ${sort.order === 'desc' ? 'desc' : 'asc'}`)
    .join(', ');
  const sortQuery = structuredSort ? ` order by ${structuredSort}` : '';
  return `${conditionQuery}${sortQuery}`;
}

function buildFilterCondition(
  filter: SourceAppConfig['filters'][number],
  storeValue: string,
  baseDate: string,
  outputRecord?: KintoneRecord
): string {
  if (!filter.field) {
    return '';
  }

  const field = fieldRef(filter.field);
  if (filter.valueFrom === 'dateRange') {
    if (!baseDate) {
      return '';
    }
    const range = calculateDateRange(filter.dateRule || 'sameDay', baseDate);
    return `${field} >= "${escapeQueryValue(range.start)}" and ${field} <= "${escapeQueryValue(range.end)}"`;
  }

  const rawValue =
    filter.valueFrom === 'store'
      ? storeValue
      : filter.valueFrom === 'baseDate'
        ? baseDate
        : filter.valueFrom === 'outputField'
          ? outputFieldValue(outputRecord, filter.outputField)
          : String(filter.value ?? '');

  if (!rawValue) {
    return '';
  }

  if (filter.operator === 'in' || filter.operator === 'not in') {
    const values = rawValue
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean)
      .map((value) => queryValue(value, filter.valueType));
    return values.length ? `${field} ${filter.operator} (${values.join(', ')})` : '';
  }

  const operator = filter.operator === 'between' ? '=' : filter.operator;
  return `${field} ${operator} ${queryValue(rawValue, filter.valueType)}`;
}

function outputFieldValue(outputRecord: KintoneRecord | undefined, fieldCode: string | undefined): string {
  if (!outputRecord || !fieldCode) {
    return '';
  }

  const value = displayValue(recordValue(outputRecord, fieldCode));
  return value == null ? '' : String(value);
}

function queryValue(value: string, valueType: 'text' | 'number' | undefined): string {
  if (valueType === 'number' && /^-?(?:\d+|\d*\.\d+)$/.test(value.trim())) {
    return value.trim();
  }
  return `"${escapeQueryValue(value)}"`;
}

export function recordValue(record: KintoneRecord, fieldCode: string): unknown {
  return record[fieldCode]?.value ?? '';
}

export function displayValue(value: unknown): string | number | boolean | Date | null {
  if (value == null) {
    return '';
  }

  if (Array.isArray(value)) {
    return value
      .map((item) => {
        if (item && typeof item === 'object' && 'name' in item) {
          return String((item as { name: unknown }).name);
        }
        if (item && typeof item === 'object' && 'code' in item) {
          return String((item as { code: unknown }).code);
        }
        return String(item);
      })
      .join(', ');
  }

  if (typeof value === 'object') {
    if ('name' in value) {
      return String((value as { name: unknown }).name);
    }
    if ('code' in value) {
      return String((value as { code: unknown }).code);
    }
    return JSON.stringify(value);
  }

  return value as string | number | boolean;
}

export function fieldDisplayValue(record: KintoneRecord, field: SourceFieldConfig): string | number | boolean | Date | null {
  const value = displayValue(recordValue(record, field.code));

  if (field.type === 'number') {
    return toNumber(value);
  }

  if (field.type === 'date' || field.type === 'datetime') {
    return toDate(value);
  }

  if (field.type === 'boolean') {
    return toBoolean(value);
  }

  return value;
}

export function fieldRef(fieldCode: string): string {
  const normalized = fieldCode.trim();
  if (!normalized || /[\s"'`()=<>!,]/.test(normalized)) {
    throw new Error(`クエリに使用できないフィールドコードです: ${fieldCode}`);
  }
  return normalized;
}

export function buildInQuery(fieldCode: string, values: string[]): string {
  const uniqueValues = Array.from(new Set(values.map((value) => value.trim()).filter(Boolean)));
  if (!uniqueValues.length) {
    return '';
  }
  return `${fieldRef(fieldCode)} in (${uniqueValues.map((value) => `"${escapeQueryValue(value)}"`).join(', ')})`;
}

function escapeQueryValue(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function toNumber(value: unknown): number | '' {
  if (value == null || value === '') {
    return '';
  }

  if (typeof value === 'number') {
    return value;
  }

  const normalized = String(value).replace(/,/g, '').trim();
  const numberValue = Number(normalized);
  return Number.isFinite(numberValue) ? numberValue : '';
}

function toDate(value: unknown): Date | '' {
  if (value == null || value === '') {
    return '';
  }

  if (value instanceof Date) {
    return value;
  }

  const dateValue = new Date(String(value));
  return Number.isNaN(dateValue.getTime()) ? '' : dateValue;
}

function toBoolean(value: unknown): boolean | '' {
  if (value == null || value === '') {
    return '';
  }

  if (typeof value === 'boolean') {
    return value;
  }

  const normalized = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'y'].includes(normalized)) {
    return true;
  }
  if (['false', '0', 'no', 'n'].includes(normalized)) {
    return false;
  }
  return '';
}

function sourceFieldValueType(kintoneType: string): SourceFieldValueType {
  if (['NUMBER', 'CALC', 'RECORD_NUMBER'].includes(kintoneType)) {
    return 'number';
  }

  if (kintoneType === 'DATE') {
    return 'date';
  }

  if (['DATETIME', 'CREATED_TIME', 'UPDATED_TIME'].includes(kintoneType)) {
    return 'datetime';
  }

  return 'text';
}
