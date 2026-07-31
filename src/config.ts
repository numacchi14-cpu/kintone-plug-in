import { defaultSources } from './shared/defaults';
import { parsePluginConfig, parseTemplateSources, serializePluginConfig } from './shared/config';
import type { BaseDateRule, PluginConfig, PluginMode } from './shared/types';

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
