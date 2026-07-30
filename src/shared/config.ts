import { defaultConfig } from './defaults';
import type {
  BaseDateRule,
  DateRangeRule,
  PluginConfig,
  PluginMode,
  SourceAppConfig,
  SourceFieldValueType,
  SourceFilterOperator,
  SourceFilterValueFrom
} from './types';

type RawPluginConfig = Record<string, string>;

const pluginModes: PluginMode[] = ['template', 'output'];
const baseDateRules: BaseDateRule[] = ['today', 'firstDayUsesYesterday', 'yesterday', 'manual'];
const sourceFieldValueTypes: SourceFieldValueType[] = ['text', 'number', 'date', 'datetime', 'boolean'];
const sourceFilterOperators: SourceFilterOperator[] = [
  '=',
  '!=',
  '>',
  '>=',
  '<',
  '<=',
  'like',
  'not like',
  'in',
  'not in',
  'between'
];
const sourceFilterValueFrom: SourceFilterValueFrom[] = ['store', 'dateRange', 'baseDate', 'fixed'];
const dateRangeRules: DateRangeRule[] = [
  'sameDay',
  'monthStartToBaseDate',
  'yearStartToBaseDate',
  'baseMonth',
  'previousMonth',
  'nextMonth',
  'baseMonthToNextMonthEnd'
];

export function parsePluginConfig(raw: RawPluginConfig | null | undefined): PluginConfig {
  const data = raw ?? {};
  const mode = pluginModes.includes(data.mode as PluginMode) ? (data.mode as PluginMode) : defaultConfig.mode;
  const baseDateRule = baseDateRules.includes(data.baseDateRule as BaseDateRule)
    ? (data.baseDateRule as BaseDateRule)
    : defaultConfig.baseDateRule;

  return {
    ...defaultConfig,
    mode,
    templateAppId: data.templateAppId ?? defaultConfig.templateAppId,
    templateReportIdField: data.templateReportIdField ?? defaultConfig.templateReportIdField,
    templateReportNameField: data.templateReportNameField ?? defaultConfig.templateReportNameField,
    templateAttachmentField: data.templateAttachmentField ?? defaultConfig.templateAttachmentField,
    templateSourcesJsonField: data.templateSourcesJsonField ?? defaultConfig.templateSourcesJsonField,
    outputReportIdField: data.outputReportIdField ?? defaultConfig.outputReportIdField,
    outputStoreField: data.outputStoreField ?? defaultConfig.outputStoreField,
    outputBaseDateField: data.outputBaseDateField ?? defaultConfig.outputBaseDateField,
    baseDateRule,
    sources: parseSources(data.sourcesJson)
  };
}

export function serializePluginConfig(config: PluginConfig): RawPluginConfig {
  return {
    mode: config.mode,
    templateAppId: config.templateAppId,
    templateReportIdField: config.templateReportIdField,
    templateReportNameField: config.templateReportNameField,
    templateAttachmentField: config.templateAttachmentField,
    templateSourcesJsonField: config.templateSourcesJsonField,
    outputReportIdField: config.outputReportIdField,
    outputStoreField: config.outputStoreField,
    outputBaseDateField: config.outputBaseDateField,
    baseDateRule: config.baseDateRule,
    sourcesJson: JSON.stringify(config.sources, null, 2)
  };
}

export function parseSources(value: string | undefined): SourceAppConfig[] {
  if (!value) {
    return defaultConfig.sources;
  }

  try {
    const parsed = JSON.parse(value) as SourceAppConfig[];
    return normalizeSources(parsed, defaultConfig.sources);
  } catch {
    return defaultConfig.sources;
  }
}

export function parseTemplateSources(value: string | undefined, fallback: SourceAppConfig[]): SourceAppConfig[] {
  if (!value?.trim()) {
    return fallback;
  }

  let parsed: SourceAppConfig[];
  try {
    parsed = JSON.parse(value) as SourceAppConfig[];
  } catch {
    throw new Error('テンプレート管理レコードの取得元アプリ設定JSONが不正です。JSON形式を確認してください。');
  }

  return normalizeSources(parsed, fallback);
}

function normalizeSources(parsed: SourceAppConfig[], fallback: SourceAppConfig[]): SourceAppConfig[] {
  if (!Array.isArray(parsed) || parsed.length === 0) {
    return fallback;
  }

  return parsed.map((source, index) => {
    const rawSource = source as unknown as Record<string, unknown>;
    const hasLegacyFilter = Array.isArray(source.filters) &&
      source.filters.some((filter) => (filter as unknown as Record<string, unknown>)?.valueFrom === 'period');
    if (
      'dateField' in rawSource ||
      'storeField' in rawSource ||
      'sort' in rawSource ||
      'periodRule' in rawSource ||
      'additionalQuery' in rawSource ||
      hasLegacyFilter
    ) {
      throw new Error('旧形式の取得元アプリ設定JSONは使用できません。簡易生成UIで設定を作り直してください。');
    }
    if (!Array.isArray(source.filters) || !Array.isArray(source.sorts)) {
      throw new Error('取得元アプリ設定JSONには filters と sorts の配列が必要です。');
    }

    return {
      key: String(source.key || `source_${index + 1}`),
      label: String(source.label || source.sheetName || `取得元${index + 1}`),
      appId: String(source.appId || ''),
      sheetName: String(source.sheetName || source.label || `元データ${index + 1}`),
      fields: Array.isArray(source.fields)
        ? source.fields
            .filter((field) => field?.code)
            .map((field) => ({
              code: String(field.code),
              label: String(field.label || field.code),
              type: sourceFieldValueTypes.includes(field.type as SourceFieldValueType)
                ? (field.type as SourceFieldValueType)
                : undefined
            }))
        : [],
      filters: source.filters
        .filter((filter) => filter?.field)
        .map((filter) => {
          const valueFrom = sourceFilterValueFrom.includes(filter.valueFrom as SourceFilterValueFrom)
            ? (filter.valueFrom as SourceFilterValueFrom)
            : 'fixed';
          return {
            field: String(filter.field),
            operator: sourceFilterOperators.includes(filter.operator as SourceFilterOperator)
              ? (filter.operator as SourceFilterOperator)
              : '=',
            valueFrom,
            value: filter.value == null ? '' : String(filter.value),
            valueType: filter.valueType === 'number' ? 'number' : 'text',
            ...(valueFrom === 'dateRange'
              ? {
                  dateRule: dateRangeRules.includes(filter.dateRule as DateRangeRule)
                    ? (filter.dateRule as DateRangeRule)
                    : 'sameDay'
                }
              : {})
          };
        }),
      sorts: source.sorts
        .filter((sort) => sort?.field)
        .map((sort) => ({
          field: String(sort.field),
          order: sort.order === 'desc' ? 'desc' : 'asc'
        }))
    };
  });
}
