import type { PluginConfig, SourceAppConfig } from './types';

export const defaultSources: SourceAppConfig[] = [
  {
    key: 'annual_budget',
    label: '年初予算',
    appId: '',
    sheetName: '年初予算',
    tableName: 'tbl_annual_budget',
    fields: [
      { code: 'store_code', label: '店舗コード' },
      { code: 'date', label: '日付', type: 'date' },
      { code: 'budget_amount', label: '予算金額', type: 'number' }
    ],
    filters: [
      { field: 'store_code', operator: '=', valueFrom: 'store', valueType: 'text' },
      {
        field: 'date',
        operator: 'between',
        valueFrom: 'dateRange',
        valueType: 'text',
        dateRule: 'yearStartToBaseDate'
      }
    ],
    sorts: [{ field: 'date', order: 'asc' }],
    lookups: []
  },
  {
    key: 'daily_plan_actual',
    label: '日別計画・実績',
    appId: '',
    sheetName: '日別計画実績',
    tableName: 'tbl_daily_plan_actual',
    fields: [
      { code: 'store_code', label: '店舗コード' },
      { code: 'date', label: '日付', type: 'date' },
      { code: 'plan_amount', label: '計画金額', type: 'number' },
      { code: 'actual_amount', label: '実績金額', type: 'number' }
    ],
    filters: [
      { field: 'store_code', operator: '=', valueFrom: 'store', valueType: 'text' },
      {
        field: 'date',
        operator: 'between',
        valueFrom: 'dateRange',
        valueType: 'text',
        dateRule: 'monthStartToBaseDate'
      }
    ],
    sorts: [{ field: 'date', order: 'asc' }],
    lookups: []
  }
];

export const defaultConfig: PluginConfig = {
  mode: 'output',
  templateAppId: '',
  templateReportIdField: 'report_id',
  templateReportNameField: 'report_name',
  templateAttachmentField: 'completed_template',
  templateSourcesJsonField: 'sources_json',
  outputReportIdField: 'report_type',
  outputStoreField: 'store',
  outputBaseDateField: 'base_date',
  outputPeriodStartField: '',
  outputPeriodEndField: '',
  outputExportedAtField: '',
  outputExporterField: '',
  outputFileNameField: '',
  outputStatusField: '',
  outputMemoField: '',
  baseDateRule: 'firstDayUsesYesterday',
  sources: defaultSources
};
