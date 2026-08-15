export type PluginMode = 'template' | 'output';

export type BaseDateRule = 'yesterday';

export type DateRangeRule =
  | 'sameDay'
  | 'previousDay'
  | 'nextDay'
  | 'baseWeek'
  | 'previousWeek'
  | 'nextWeek'
  | 'monthStartToBaseDate'
  | 'yearStartToBaseDate'
  | 'baseMonthStart'
  | 'baseMonthEnd'
  | 'baseMonth'
  | 'previousMonthStart'
  | 'previousMonthEnd'
  | 'previousMonth'
  | 'nextMonthStart'
  | 'nextMonthEnd'
  | 'nextMonth'
  | 'baseMonthToNextMonthEnd'
  | 'baseYearStart'
  | 'baseYearEnd'
  | 'baseYear'
  | 'baseFirstHalf'
  | 'baseSecondHalf'
  | 'previousYearStart'
  | 'previousYearEnd'
  | 'previousYear'
  | 'nextYearStart'
  | 'nextYearEnd'
  | 'nextYear'
  | 'sameDayPreviousYear'
  | 'sameMonthPreviousYear'
  | 'previousMonthPreviousYear';

export type SourceFieldValueType = 'text' | 'number' | 'date' | 'datetime' | 'boolean';

export type SourceFilterOperator = '=' | '!=' | '>' | '>=' | '<' | '<=' | 'like' | 'not like' | 'in' | 'not in' | 'between';

export type SourceFilterValueFrom = 'store' | 'dateRange' | 'baseDate' | 'outputField' | 'fixed';

export interface SourceFilterConfig {
  field: string;
  operator: SourceFilterOperator;
  valueFrom: SourceFilterValueFrom;
  value?: string;
  outputField?: string;
  valueType?: 'text' | 'number';
  dateRule?: DateRangeRule;
}

export interface SourceSortConfig {
  field: string;
  order: 'asc' | 'desc';
}

export interface SourceLookupConfig {
  sourceField: string;
  masterAppId: string;
  masterKeyField: string;
  masterFields: SourceFieldConfig[];
}

export interface SourceFieldConfig {
  code: string;
  label: string;
  type?: SourceFieldValueType;
}

export interface SourceAppConfig {
  key: string;
  label: string;
  appId: string;
  sheetName: string;
  tableName?: string;
  fields: SourceFieldConfig[];
  filters: SourceFilterConfig[];
  sorts: SourceSortConfig[];
  lookups: SourceLookupConfig[];
}

export interface PluginConfig {
  mode: PluginMode;
  templateAppId: string;
  templateReportIdField: string;
  templateReportNameField: string;
  templateAttachmentField: string;
  templateSourcesJsonField: string;
  outputAppId: string;
  outputReportIdField: string;
  outputStoreField: string;
  outputBaseDateField: string;
  outputPeriodStartField: string;
  outputPeriodEndField: string;
  outputExportedAtField: string;
  outputExporterField: string;
  outputFileNameField: string;
  outputStatusField: string;
  outputMemoField: string;
  baseDateRule: BaseDateRule;
  sources: SourceAppConfig[];
}

export interface ReportContext {
  reportId: string;
  reportName: string;
  store: string;
  baseDate: string;
  periodStart: string;
  periodEnd: string;
  exportedAt: string;
  exporter: string;
}

export interface SourceRows {
  source: SourceAppConfig;
  rows: Record<string, unknown>[];
  periodStart: string;
  periodEnd: string;
  query: string;
}

export type KintoneRecord = Record<string, { type: string; value: unknown }>;
