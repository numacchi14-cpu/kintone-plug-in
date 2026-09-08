import { defaultSources } from './shared/defaults';
import { parsePluginConfig, parseTemplateSources, serializePluginConfig } from './shared/config';
import type { BaseDateRule, PickerDisplayField, PluginConfig, PluginMode } from './shared/types';

const pickerDisplayFieldIds: PickerDisplayField[] = ['reportType', 'reportName', 'store', 'yesterdayBase', 'recordBase'];

declare const kintone: any;

const pluginId = kintone.$PLUGIN_ID;
const config = parsePluginConfig(kintone.plugin.app.getConfig(pluginId));

setInputValue('mode', config.mode);
setInputValue('templateAppId', config.templateAppId);
setInputValue('templateReportIdField', config.templateReportIdField);
setInputValue('templateReportNameField', config.templateReportNameField);
setInputValue('templateAttachmentField', config.templateAttachmentField);
setInputValue('templateSourcesJsonField', config.templateSourcesJsonField);
setInputValue('outputAppId', config.outputAppId);
setInputValue('outputReportIdField', config.outputReportIdField);
setInputValue('outputReportNameField', config.outputReportNameField);
setInputValue('outputStoreField', config.outputStoreField);
setInputValue('outputBaseDateField', config.outputBaseDateField);
setInputValue('outputPeriodStartField', config.outputPeriodStartField);
setInputValue('outputPeriodEndField', config.outputPeriodEndField);
setInputValue('outputExportedAtField', config.outputExportedAtField);
setInputValue('outputExporterField', config.outputExporterField);
setInputValue('outputFileNameField', config.outputFileNameField);
setInputValue('outputStatusField', config.outputStatusField);
setInputValue('outputMemoField', config.outputMemoField);
setInputValue('baseDateRule', 'yesterday');
for (const field of pickerDisplayFieldIds) {
  setCheckboxValue(`pickerDisplayField_${field}`, config.pickerDisplayFields.includes(field));
}
setInputValue('sourcesJson', JSON.stringify(config.sources.length ? config.sources : defaultSources, null, 2));

document.getElementById('save')?.addEventListener('click', () => {
  try {
    const nextConfig: PluginConfig = {
      mode: getInputValue('mode') as PluginMode,
      templateAppId: getInputValue('templateAppId'),
      templateReportIdField: getInputValue('templateReportIdField'),
      templateReportNameField: getInputValue('templateReportNameField'),
      templateAttachmentField: getInputValue('templateAttachmentField'),
      templateSourcesJsonField: getInputValue('templateSourcesJsonField'),
      outputAppId: getInputValue('outputAppId'),
      outputReportIdField: getInputValue('outputReportIdField'),
      outputReportNameField: getInputValue('outputReportNameField'),
      outputStoreField: getInputValue('outputStoreField'),
      outputBaseDateField: getInputValue('outputBaseDateField'),
      outputPeriodStartField: getInputValue('outputPeriodStartField'),
      outputPeriodEndField: getInputValue('outputPeriodEndField'),
      outputExportedAtField: getInputValue('outputExportedAtField'),
      outputExporterField: getInputValue('outputExporterField'),
      outputFileNameField: getInputValue('outputFileNameField'),
      outputStatusField: getInputValue('outputStatusField'),
      outputMemoField: getInputValue('outputMemoField'),
      baseDateRule: 'yesterday' as BaseDateRule,
      pickerDisplayFields: pickerDisplayFieldIds.filter((field) => getCheckboxValue(`pickerDisplayField_${field}`)),
      sources: parseTemplateSources(getInputValue('sourcesJson'), defaultSources)
    };

    kintone.plugin.app.setConfig(serializePluginConfig(nextConfig), () => {
      window.location.href = '../../flow?app=' + kintone.app.getId();
    });
  } catch (error) {
    window.alert(error instanceof Error ? error.message : '設定の保存に失敗しました。');
  }
});

document.getElementById('cancel')?.addEventListener('click', () => {
  window.location.href = '../../' + kintone.app.getId() + '/plugin/';
});

function setInputValue(id: string, value: string): void {
  const element = document.getElementById(id);
  if (element instanceof HTMLInputElement || element instanceof HTMLSelectElement || element instanceof HTMLTextAreaElement) {
    element.value = value;
  }
}

function getInputValue(id: string): string {
  const element = document.getElementById(id);
  if (element instanceof HTMLInputElement || element instanceof HTMLSelectElement || element instanceof HTMLTextAreaElement) {
    return element.value.trim();
  }
  return '';
}

function setCheckboxValue(id: string, checked: boolean): void {
  const element = document.getElementById(id);
  if (element instanceof HTMLInputElement) {
    element.checked = checked;
  }
}

function getCheckboxValue(id: string): boolean {
  const element = document.getElementById(id);
  return element instanceof HTMLInputElement && element.checked;
}
