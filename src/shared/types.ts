export type PluginMode = 'template' | 'output';

export type BaseDateRule = 'today' | 'firstDayUsesYesterday' | 'yesterday' | 'manual';

export type DateRangeRule =
  | 'sameDay'
  | 'monthStartToBaseDate'
  | 'yearStartToBaseDate'
  | 'baseMonth'
  | 'previousMonth'
  | 'nextMonth'
  | 'baseMonthToNextMonthEnd';

export type SourceFieldValueType = 'text' | 'number' | 'date' | 'datetime' | 'boolean';

export type SourceFilterOperator = '=' | '!=' | '>' | '>=' | '<' | '<=' | 'like' | 'not like' | 'in' | 'not in' | 'between';

export type SourceFilterValueFrom = 'store' | 'dateRange' | 'baseDate' | 'fixed';

export interface SourceFilterConfig {
  field: string;
  operator: SourceFilterOperator;
  valueFrom: SourceFilterValueFrom;
  value?: string;
  valueType?: 'text' | 'number';
  dateRule?: DateRangeRule;
}

export interface SourceSortConfig {
  field: string;
  order: 'asc' | 'desc';
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
  fields: SourceFieldConfig[];
  filters: SourceFilterConfig[];
  sorts: SourceSortConfig[];
}

export interface PluginConfig {
  mode: PluginMode;
  templateAppId: string;
  templateReportIdField: string;
  templateReportNameField: string;
  templateAttachmentField: string;
  templateSourcesJsonField: string;
  outputReportIdField: string;
  outputStoreField: string;
  outputBaseDateField: string;
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
}

export type KintoneRecord = Record<string, { type: string; value: unknown }>;
