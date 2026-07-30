import { parsePluginConfig, parseTemplateSources } from './shared/config';
import { calculateDateRange, formatDate, resolveBaseDate } from './shared/dateRules';
import { buildFileName, createInitialTemplate, downloadWorkbook, fillReportTemplate } from './shared/excel';
import {
  buildSourceQuery,
  downloadKintoneFile,
  fieldDisplayValue,
  findTemplateRecord,
  getAllRecords,
  getSourceAppFields,
  recordValue
} from './shared/kintoneApi';
import type {
  KintoneRecord,
  PluginConfig,
  ReportContext,
  SourceAppConfig,
  SourceFieldConfig,
  SourceFilterConfig,
  SourceSortConfig,
  SourceRows
} from './shared/types';

declare const kintone: any;

const pluginId = kintone.$PLUGIN_ID;
const pluginConfig = parsePluginConfig(kintone.plugin.app.getConfig(pluginId));
let sourceBlockSequence = 0;

kintone.events.on(['app.record.create.show'], (event: any) => {
  if (pluginConfig.mode !== 'output') {
    return event;
  }

  const field = event.record[pluginConfig.outputBaseDateField];
  if (field && !field.value && pluginConfig.baseDateRule !== 'manual') {
    field.value = resolveBaseDate(pluginConfig.baseDateRule);
  }

  return event;
});

kintone.events.on(['app.record.create.show', 'app.record.edit.show'], (event: any) => {
  if (pluginConfig.mode === 'template') {
    renderSourceJsonBuilder(pluginConfig);
  } else {
    renderBaseDateUpdateButton(pluginConfig);
  }

  return event;
});

kintone.events.on(['app.record.detail.show'], (event: any) => {
  const header = kintone.app.record.getHeaderMenuSpaceElement();
  if (!header || header.querySelector('[data-krp-toolbar="true"]')) {
    return event;
  }

  const toolbar = document.createElement('span');
  toolbar.className = 'krp-toolbar';
  toolbar.dataset.krpToolbar = 'true';

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'krp-button';
  button.textContent = pluginConfig.mode === 'template' ? '初期Excelテンプレート作成' : 'Excel帳票出力';

  const status = document.createElement('span');
  status.className = 'krp-status';
  status.textContent = pluginConfig.mode === 'template' ? 'テンプレート管理モード' : '帳票出力モード';

  button.addEventListener('click', async () => {
    await runWithStatus(button, status, async () => {
      if (pluginConfig.mode === 'template') {
        await generateInitialTemplate(pluginConfig, event.record);
      } else {
        await exportReport(pluginConfig, event.record);
      }
    });
  });

  toolbar.append(button, status);
  header.appendChild(toolbar);
  return event;
});

async function generateInitialTemplate(config: PluginConfig, record: KintoneRecord): Promise<void> {
  const reportName = String(recordValue(record, config.templateReportNameField) || 'Excel帳票テンプレート');
  const sources = resolveTemplateSources(config, record);
  const buffer = await createInitialTemplate({ ...config, sources }, reportName);
  downloadWorkbook(buffer, `${reportName.replace(/[\\/:*?"<>|]/g, '_')}_初期テンプレート.xlsx`);
}

function renderBaseDateUpdateButton(config: PluginConfig): void {
  if (config.baseDateRule === 'manual') {
    return;
  }

  const header = kintone.app.record.getHeaderMenuSpaceElement();
  if (!header || header.querySelector('[data-krp-base-date-tool="true"]')) {
    return;
  }

  const tool = document.createElement('span');
  tool.className = 'krp-toolbar';
  tool.dataset.krpBaseDateTool = 'true';

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'krp-button krp-button--secondary';
  button.textContent = '基準日を自動設定';
  button.title = 'プラグイン設定の初期値ルールで基準日を再設定します';
  button.addEventListener('click', () => {
    const current = kintone.app.record.get();
    const field = current?.record?.[config.outputBaseDateField];
    if (!field) {
      window.alert(`基準日フィールド「${config.outputBaseDateField}」がフォームにありません。`);
      return;
    }

    field.value = resolveBaseDate(config.baseDateRule, new Date(), '');
    kintone.app.record.set(current);
  });

  tool.appendChild(button);
  header.appendChild(tool);
}

function renderSourceJsonBuilder(config: PluginConfig): void {
  const header = kintone.app.record.getHeaderMenuSpaceElement();
  if (!header || header.querySelector('[data-krp-source-builder="true"]')) {
    return;
  }

  const panel = document.createElement('details');
  panel.className = 'krp-builder';
  panel.dataset.krpSourceBuilder = 'true';
  panel.open = false;
  panel.innerHTML = `
    <summary>取得元アプリ設定JSONをかんたん作成</summary>
    <div class="krp-builder__actions">
      <button type="button" data-krp-action="add-source">取得元アプリを追加</button>
      <button type="button" data-krp-action="apply-json">JSONをフィールドへ反映</button>
      <button type="button" data-krp-action="load-json">現在のJSONを読み込む</button>
    </div>
    <div data-krp-builder="sources"></div>
    <p class="krp-builder__note">反映後は、このレコードを保存してください。取得元アプリは複数追加できます。列順は各ブロック内の出力フィールドの行順です。</p>
  `;

  const addSourceButton = panel.querySelector('[data-krp-action="add-source"]');
  const applyButton = panel.querySelector('[data-krp-action="apply-json"]');
  const loadButton = panel.querySelector('[data-krp-action="load-json"]');

  addSourceButton?.addEventListener('click', () => addBuilderSourceBlock(panel));
  applyButton?.addEventListener('click', () => applyBuilderJson(config, panel));
  loadButton?.addEventListener('click', () => loadBuilderJson(config, panel));

  header.appendChild(panel);
  addBuilderSourceBlock(panel);
}

function addBuilderSourceBlock(panel: HTMLElement, source?: SourceAppConfig): void {
  const sourcesContainer = panel.querySelector('[data-krp-builder="sources"]');
  if (!sourcesContainer) {
    return;
  }

  const sourceNumber = sourcesContainer.querySelectorAll('.krp-builder__source').length + 1;
  const sourceBlockId = `krp-source-fields-${++sourceBlockSequence}`;
  const sourceBlock = document.createElement('details');
  sourceBlock.className = 'krp-builder__source';
  sourceBlock.open = true;
  sourceBlock.innerHTML = `
    <summary>${escapeHtml(source?.label || source?.sheetName || `取得元アプリ ${sourceNumber}`)}</summary>
    <div class="krp-builder__source-actions">
      <button type="button" data-krp-action="import-fields">フィールド一覧を取得</button>
      <button type="button" data-krp-action="duplicate-source">複製</button>
      <button type="button" data-krp-action="remove-source">削除</button>
      <span class="krp-builder__source-status" data-krp-source="importStatus"></span>
    </div>
    <div class="krp-builder__grid">
      <label>名前<input data-krp-source="label" placeholder="例: 実績" value="${escapeHtml(source?.label ?? '')}"></label>
      <label>アプリID<input data-krp-source="appId" inputmode="numeric" placeholder="例: 99" value="${escapeHtml(source?.appId ?? '')}"></label>
      <label>Excelシート名<input data-krp-source="sheetName" placeholder="例: 日別計画実績" value="${escapeHtml(source?.sheetName ?? '')}"></label>
    </div>
    <div class="krp-builder__rules">
      <div class="krp-builder__rules-head">
        <strong>抽出条件</strong>
        <button type="button" data-krp-action="add-filter">条件を追加</button>
      </div>
      <div data-krp-source="filters"></div>
      <p class="krp-builder__note">条件がない場合は全件取得します。日付条件ごとに、基準日から計算する期間を選べます。</p>
    </div>
    <div class="krp-builder__rules">
      <div class="krp-builder__rules-head">
        <strong>並び順</strong>
        <button type="button" data-krp-action="add-sort">並び順を追加</button>
      </div>
      <div data-krp-source="sorts"></div>
    </div>
    <div class="krp-builder__fields">
      <div class="krp-builder__fields-head">
        <strong>出力フィールド</strong>
        <button type="button" data-krp-action="add-field">行を追加</button>
        <button type="button" data-krp-action="select-all-fields">全選択</button>
        <button type="button" data-krp-action="clear-all-fields">全解除</button>
        <button type="button" data-krp-action="remove-selected-fields">選択行を削除</button>
      </div>
      <div data-krp-source="fields"></div>
      <datalist id="${sourceBlockId}" data-krp-source="fieldCodes"></datalist>
      <p class="krp-builder__note">チェックした行だけJSONへ反映されます。行の上下ボタンでExcelの列順を変更できます。</p>
    </div>
  `;

  sourceBlock.querySelector('[data-krp-action="import-fields"]')?.addEventListener('click', async () => {
    await importBuilderFields(sourceBlock);
  });
  sourceBlock.querySelector('[data-krp-action="add-filter"]')?.addEventListener('click', () => addBuilderFilterRow(sourceBlock));
  sourceBlock.querySelector('[data-krp-action="add-sort"]')?.addEventListener('click', () => addBuilderSortRow(sourceBlock));
  sourceBlock.querySelector('[data-krp-action="add-field"]')?.addEventListener('click', () => addBuilderFieldRow(sourceBlock));
  sourceBlock.querySelector('[data-krp-action="select-all-fields"]')?.addEventListener('click', () => {
    setBuilderFieldSelection(sourceBlock, true);
  });
  sourceBlock.querySelector('[data-krp-action="clear-all-fields"]')?.addEventListener('click', () => {
    setBuilderFieldSelection(sourceBlock, false);
  });
  sourceBlock.querySelector('[data-krp-action="remove-selected-fields"]')?.addEventListener('click', () => {
    removeSelectedBuilderFields(sourceBlock);
  });
  sourceBlock.querySelector('[data-krp-action="duplicate-source"]')?.addEventListener('click', () => {
    addBuilderSourceBlock(panel, readBuilderSource(sourceBlock, sourceNumber));
  });
  sourceBlock.querySelector('[data-krp-action="remove-source"]')?.addEventListener('click', () => {
    if (panel.querySelectorAll('.krp-builder__source').length <= 1) {
      window.alert('取得元アプリは最低1つ必要です。');
      return;
    }
    sourceBlock.remove();
    refreshSourceSummaries(panel);
  });

  sourcesContainer.appendChild(sourceBlock);

  const initialFilters = builderFilters(source);
  initialFilters.forEach((filter) => addBuilderFilterRow(sourceBlock, filter));
  builderSorts(source).forEach((sort) => addBuilderSortRow(sourceBlock, sort));

  if (source?.fields.length) {
    source.fields.forEach((field) => addBuilderFieldRow(sourceBlock, field));
  } else {
    addBuilderFieldRow(sourceBlock, { code: '日付', label: '日付', type: 'date' });
    addBuilderFieldRow(sourceBlock, { code: '店舗名', label: '店舗名', type: 'text' });
  }
  refreshFieldCodeSuggestions(sourceBlock);

  Array.from(sourceBlock.querySelectorAll('[data-krp-source="label"], [data-krp-source="sheetName"]')).forEach((input) => {
    input.addEventListener('input', () => refreshSourceSummaries(panel));
  });
  refreshSourceSummaries(panel);
}

function addBuilderFilterRow(sourceBlock: HTMLElement, filter: Partial<SourceFilterConfig> = {}): void {
  const filters = sourceBlock.querySelector('[data-krp-source="filters"]');
  if (!filters) {
    return;
  }

  const row = document.createElement('div');
  row.className = 'krp-builder__filter-row';
  row.innerHTML = `
    <input data-krp-filter="field" list="${fieldCodeListId(sourceBlock)}" placeholder="フィールドコード" value="${escapeHtml(filter.field ?? '')}">
    <select data-krp-filter="operator">
      <option value="=">等しい</option>
      <option value="!=">等しくない</option>
      <option value=">">より大きい</option>
      <option value=">=">以上</option>
      <option value="<">より小さい</option>
      <option value="<=">以下</option>
      <option value="like">含む</option>
      <option value="not like">含まない</option>
      <option value="in">いずれか</option>
      <option value="not in">いずれでもない</option>
      <option value="between">対象期間内</option>
    </select>
    <select data-krp-filter="valueFrom">
      <option value="store">出力アプリの対象店舗</option>
      <option value="dateRange">基準日から期間計算</option>
      <option value="baseDate">出力アプリの基準日</option>
      <option value="fixed">固定値</option>
    </select>
    <select data-krp-filter="dateRule">
      <option value="sameDay">基準日と同じ日</option>
      <option value="monthStartToBaseDate">月初から基準日</option>
      <option value="yearStartToBaseDate">1月1日から基準日</option>
      <option value="baseMonth">基準日の月全体</option>
      <option value="previousMonth">基準日の前月</option>
      <option value="nextMonth">基準日の翌月</option>
      <option value="baseMonthToNextMonthEnd">月初から翌月末</option>
    </select>
    <input data-krp-filter="value" placeholder="固定値。複数はカンマ区切り" value="${escapeHtml(filter.value ?? '')}">
    <select data-krp-filter="valueType">
      <option value="text">文字・日付</option>
      <option value="number">数値</option>
    </select>
    <button type="button" data-krp-action="remove-filter">削除</button>
  `;

  const operator = row.querySelector('[data-krp-filter="operator"]');
  const valueFrom = row.querySelector('[data-krp-filter="valueFrom"]');
  const dateRule = row.querySelector('[data-krp-filter="dateRule"]');
  const valueType = row.querySelector('[data-krp-filter="valueType"]');
  if (operator instanceof HTMLSelectElement) {
    operator.value = filter.operator ?? '=';
  }
  if (valueFrom instanceof HTMLSelectElement) {
    valueFrom.value = filter.valueFrom ?? 'fixed';
  }
  if (dateRule instanceof HTMLSelectElement) {
    dateRule.value = filter.dateRule ?? 'monthStartToBaseDate';
  }
  if (valueType instanceof HTMLSelectElement) {
    valueType.value = filter.valueType ?? 'text';
  }

  valueFrom?.addEventListener('change', () => syncBuilderFilterRow(row));
  row.querySelector('[data-krp-action="remove-filter"]')?.addEventListener('click', () => row.remove());
  filters.appendChild(row);
  syncBuilderFilterRow(row);
}

function syncBuilderFilterRow(row: HTMLElement): void {
  const operator = row.querySelector('[data-krp-filter="operator"]');
  const valueFrom = row.querySelector('[data-krp-filter="valueFrom"]');
  const dateRule = row.querySelector('[data-krp-filter="dateRule"]');
  const value = row.querySelector('[data-krp-filter="value"]');

  if (!(operator instanceof HTMLSelectElement) || !(valueFrom instanceof HTMLSelectElement)) {
    return;
  }

  const usesDateRange = valueFrom.value === 'dateRange';
  if (usesDateRange) {
    operator.value = 'between';
  } else if (operator.value === 'between') {
    operator.value = '=';
  }
  operator.disabled = usesDateRange;

  if (dateRule instanceof HTMLSelectElement) {
    dateRule.hidden = !usesDateRange;
    dateRule.disabled = !usesDateRange;
  }

  if (value instanceof HTMLInputElement) {
    value.disabled = valueFrom.value !== 'fixed';
  }
}

function addBuilderSortRow(sourceBlock: HTMLElement, sort: Partial<SourceSortConfig> = {}): void {
  const sorts = sourceBlock.querySelector('[data-krp-source="sorts"]');
  if (!sorts) {
    return;
  }

  const row = document.createElement('div');
  row.className = 'krp-builder__sort-row';
  row.innerHTML = `
    <input data-krp-sort="field" list="${fieldCodeListId(sourceBlock)}" placeholder="フィールドコード" value="${escapeHtml(sort.field ?? '')}">
    <select data-krp-sort="order">
      <option value="asc">昇順</option>
      <option value="desc">降順</option>
    </select>
    <div class="krp-builder__field-actions">
      <button type="button" data-krp-action="move-sort-up" aria-label="上へ移動" title="上へ移動">↑</button>
      <button type="button" data-krp-action="move-sort-down" aria-label="下へ移動" title="下へ移動">↓</button>
      <button type="button" data-krp-action="remove-sort">削除</button>
    </div>
  `;

  const order = row.querySelector('[data-krp-sort="order"]');
  if (order instanceof HTMLSelectElement) {
    order.value = sort.order ?? 'asc';
  }

  row.querySelector('[data-krp-action="move-sort-up"]')?.addEventListener('click', () => {
    row.previousElementSibling?.before(row);
  });
  row.querySelector('[data-krp-action="move-sort-down"]')?.addEventListener('click', () => {
    row.nextElementSibling?.after(row);
  });
  row.querySelector('[data-krp-action="remove-sort"]')?.addEventListener('click', () => row.remove());
  sorts.appendChild(row);
}

function builderFilters(source?: SourceAppConfig): SourceFilterConfig[] {
  if (source) {
    return source.filters;
  }

  return [
    { field: '店舗名', operator: '=', valueFrom: 'store', valueType: 'text' },
    {
      field: '日付',
      operator: 'between',
      valueFrom: 'dateRange',
      valueType: 'text',
      dateRule: 'monthStartToBaseDate'
    }
  ];
}

function builderSorts(source?: SourceAppConfig): SourceSortConfig[] {
  return source?.sorts ?? [];
}

function fieldCodeListId(sourceBlock: Element): string {
  return sourceBlock.querySelector('[data-krp-source="fieldCodes"]')?.id || '';
}

function addBuilderFieldRow(
  sourceBlock: HTMLElement,
  field: { code?: string; label?: string; type?: string } = {},
  selected = true
): void {
  const fields = sourceBlock.querySelector('[data-krp-source="fields"]');
  if (!fields) {
    return;
  }

  const row = document.createElement('div');
  row.className = 'krp-builder__field-row';
  row.innerHTML = `
    <input type="checkbox" data-krp-field="selected" aria-label="出力対象" title="JSONへ反映する"${selected ? ' checked' : ''}>
    <input data-krp-field="code" placeholder="フィールドコード" value="${escapeHtml(field.code ?? '')}">
    <input data-krp-field="label" placeholder="Excel見出し" value="${escapeHtml(field.label ?? '')}">
    <select data-krp-field="type">
      <option value="text">文字列</option>
      <option value="number">数値</option>
      <option value="date">日付</option>
      <option value="datetime">日時</option>
      <option value="boolean">真偽値</option>
    </select>
    <div class="krp-builder__field-actions">
      <button type="button" data-krp-action="move-field-up" aria-label="上へ移動" title="上へ移動">↑</button>
      <button type="button" data-krp-action="move-field-down" aria-label="下へ移動" title="下へ移動">↓</button>
      <button type="button" data-krp-action="remove-field">削除</button>
    </div>
  `;

  const typeSelect = row.querySelector('[data-krp-field="type"]');
  if (typeSelect instanceof HTMLSelectElement) {
    typeSelect.value = field.type ?? 'text';
  }

  row.querySelector('[data-krp-field="code"]')?.addEventListener('input', () => refreshFieldCodeSuggestions(sourceBlock));
  row.querySelector('[data-krp-action="move-field-up"]')?.addEventListener('click', () => {
    row.previousElementSibling?.before(row);
  });
  row.querySelector('[data-krp-action="move-field-down"]')?.addEventListener('click', () => {
    row.nextElementSibling?.after(row);
  });
  row.querySelector('[data-krp-action="remove-field"]')?.addEventListener('click', () => {
    row.remove();
    refreshFieldCodeSuggestions(sourceBlock);
  });
  fields.appendChild(row);
  refreshFieldCodeSuggestions(sourceBlock);
}

async function importBuilderFields(sourceBlock: HTMLElement): Promise<void> {
  const appId = getSourceValue(sourceBlock, 'appId');
  const importButton = sourceBlock.querySelector('[data-krp-action="import-fields"]');
  const status = sourceBlock.querySelector('[data-krp-source="importStatus"]');

  if (!appId) {
    window.alert('フィールド一覧を取得するアプリIDを入力してください。');
    return;
  }

  if (importButton instanceof HTMLButtonElement) {
    importButton.disabled = true;
  }
  if (status) {
    status.textContent = '取得中...';
  }

  try {
    const currentFields = readBuilderFieldRows(sourceBlock);
    const currentByCode = new Map(currentFields.map((field) => [field.code, field]));
    const result = await getSourceAppFields(appId);
    const fieldsContainer = sourceBlock.querySelector('[data-krp-source="fields"]');

    if (!result.fields.length) {
      throw new Error('出力に対応したフィールドが見つかりません。');
    }

    if (fieldsContainer) {
      fieldsContainer.innerHTML = '';
      result.fields.forEach((field) => {
        const current = currentByCode.get(field.code);
        addBuilderFieldRow(
          sourceBlock,
          {
            code: field.code,
            label: current?.label || field.label,
            type: field.type
          },
          current?.selected ?? true
        );
      });
    }

    refreshFieldCodeSuggestions(sourceBlock);
    if (status) {
      const skipped = result.skippedCount ? `、非対応${result.skippedCount}件を除外` : '';
      status.textContent = `${result.fields.length}件を取得${skipped}`;
    }
  } catch (error) {
    const message = errorMessage(error, 'フィールド一覧を取得できませんでした。');
    if (status) {
      status.textContent = '取得エラー';
    }
    window.alert(`アプリID ${appId} のフィールド一覧を取得できませんでした。\n${message}`);
  } finally {
    if (importButton instanceof HTMLButtonElement) {
      importButton.disabled = false;
    }
  }
}

function readBuilderFieldRows(
  sourceBlock: Element
): Array<SourceFieldConfig & { selected: boolean }> {
  return Array.from(sourceBlock.querySelectorAll('.krp-builder__field-row'))
    .map((row) => {
      const code = getRowValue(row, 'code');
      const label = getRowValue(row, 'label') || code;
      const type = getRowValue(row, 'type') as SourceFieldConfig['type'];
      const selectedElement = row.querySelector('[data-krp-field="selected"]');
      return {
        code,
        label,
        type,
        selected: selectedElement instanceof HTMLInputElement ? selectedElement.checked : true
      };
    })
    .filter((field) => field.code);
}

function setBuilderFieldSelection(sourceBlock: Element, selected: boolean): void {
  sourceBlock.querySelectorAll('[data-krp-field="selected"]').forEach((element) => {
    if (element instanceof HTMLInputElement) {
      element.checked = selected;
    }
  });
}

function removeSelectedBuilderFields(sourceBlock: HTMLElement): void {
  sourceBlock.querySelectorAll('.krp-builder__field-row').forEach((row) => {
    const checkbox = row.querySelector('[data-krp-field="selected"]');
    if (checkbox instanceof HTMLInputElement && checkbox.checked) {
      row.remove();
    }
  });
  refreshFieldCodeSuggestions(sourceBlock);
}

function refreshFieldCodeSuggestions(sourceBlock: Element): void {
  const datalist = sourceBlock.querySelector('[data-krp-source="fieldCodes"]');
  if (!(datalist instanceof HTMLDataListElement)) {
    return;
  }

  const codes = readBuilderFieldRows(sourceBlock).map((field) => field.code);
  datalist.replaceChildren(
    ...codes.map((code) => {
      const option = document.createElement('option');
      option.value = code;
      return option;
    })
  );
}

function applyBuilderJson(config: PluginConfig, panel: HTMLElement): void {
  const record = kintone.app.record.get();
  const sources = Array.from(panel.querySelectorAll('.krp-builder__source')).map((sourceBlock, index) =>
    readBuilderSource(sourceBlock, index + 1)
  );

  const invalidSource = sources.find((source) => !source.appId || !source.sheetName || !source.fields.length);
  if (invalidSource) {
    window.alert(`取得元アプリ「${invalidSource.label || invalidSource.key}」のアプリID、Excelシート名、出力フィールドを入力してください。`);
    return;
  }

  const field = record.record[config.templateSourcesJsonField];
  if (!field) {
    window.alert(`取得元アプリ設定JSONフィールド「${config.templateSourcesJsonField}」がフォームにありません。`);
    return;
  }

  field.value = JSON.stringify(sources, null, 2);
  kintone.app.record.set(record);
  window.alert('取得元アプリ設定JSONへ反映しました。レコードを保存してください。');
}

function loadBuilderJson(config: PluginConfig, panel: HTMLElement): void {
  const record = kintone.app.record.get();
  const value = String(record.record[config.templateSourcesJsonField]?.value || '');

  if (!value.trim()) {
    window.alert('現在の取得元アプリ設定JSONは空です。');
    return;
  }

  try {
    const sources = parseTemplateSources(value, []);
    if (!sources.length) {
      window.alert('読み込める取得元アプリ設定がありません。');
      return;
    }

    const sourcesContainer = panel.querySelector('[data-krp-builder="sources"]');
    if (sourcesContainer) {
      sourcesContainer.innerHTML = '';
      sources.forEach((source) => addBuilderSourceBlock(panel, source));
    }
  } catch (error) {
    window.alert(error instanceof Error ? error.message : '取得元アプリ設定JSONを読み込めません。');
  }
}

function readBuilderSource(sourceBlock: Element, sourceNumber: number): SourceAppConfig {
  const label = getSourceValue(sourceBlock, 'label');
  const fields: SourceAppConfig['fields'] = [];
  const filters: SourceFilterConfig[] = [];
  const sorts: SourceSortConfig[] = [];

  Array.from(sourceBlock.querySelectorAll('.krp-builder__field-row')).forEach((row) => {
    const selected = row.querySelector('[data-krp-field="selected"]');
    if (selected instanceof HTMLInputElement && !selected.checked) {
      return;
    }

    const code = getRowValue(row, 'code');
    if (!code) {
      return;
    }

    const fieldLabel = getRowValue(row, 'label') || code;
    const type = getRowValue(row, 'type');
    fields.push({
      code,
      label: fieldLabel,
      ...(type === 'text' ? {} : { type: type as SourceAppConfig['fields'][number]['type'] })
    });
  });

  sourceBlock.querySelectorAll('.krp-builder__filter-row').forEach((row) => {
    const field = getBuilderControlValue(row, 'filter', 'field');
    if (!field) {
      return;
    }
    const valueFrom = getBuilderControlValue(row, 'filter', 'valueFrom') as SourceFilterConfig['valueFrom'];
    const filter: SourceFilterConfig = {
      field,
      operator: getBuilderControlValue(row, 'filter', 'operator') as SourceFilterConfig['operator'],
      valueFrom,
      value: getBuilderControlValue(row, 'filter', 'value'),
      valueType: getBuilderControlValue(row, 'filter', 'valueType') === 'number' ? 'number' : 'text'
    };
    if (valueFrom === 'dateRange') {
      filter.dateRule = getBuilderControlValue(row, 'filter', 'dateRule') as SourceFilterConfig['dateRule'];
    }
    filters.push(filter);
  });

  sourceBlock.querySelectorAll('.krp-builder__sort-row').forEach((row) => {
    const field = getBuilderControlValue(row, 'sort', 'field');
    if (!field) {
      return;
    }
    sorts.push({
      field,
      order: getBuilderControlValue(row, 'sort', 'order') === 'desc' ? 'desc' : 'asc'
    });
  });

  return {
    key: toSourceKey(label || getSourceValue(sourceBlock, 'sheetName') || `source_${sourceNumber}`),
    label,
    appId: getSourceValue(sourceBlock, 'appId'),
    sheetName: getSourceValue(sourceBlock, 'sheetName'),
    fields,
    filters,
    sorts
  };
}

function getSourceValue(sourceBlock: Element, name: string): string {
  const element = sourceBlock.querySelector(`[data-krp-source="${name}"]`);
  return element instanceof HTMLInputElement || element instanceof HTMLSelectElement ? element.value.trim() : '';
}

function getRowValue(row: Element, name: string): string {
  const element = row.querySelector(`[data-krp-field="${name}"]`);
  return element instanceof HTMLInputElement || element instanceof HTMLSelectElement ? element.value.trim() : '';
}

function getBuilderControlValue(row: Element, group: 'filter' | 'sort', name: string): string {
  const element = row.querySelector(`[data-krp-${group}="${name}"]`);
  return element instanceof HTMLInputElement || element instanceof HTMLSelectElement ? element.value.trim() : '';
}

function refreshSourceSummaries(panel: HTMLElement): void {
  Array.from(panel.querySelectorAll('.krp-builder__source')).forEach((sourceBlock, index) => {
    const summary = sourceBlock.querySelector('summary');
    if (summary) {
      summary.textContent =
        getSourceValue(sourceBlock, 'label') || getSourceValue(sourceBlock, 'sheetName') || `取得元アプリ ${index + 1}`;
    }
  });
}

function toSourceKey(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'source';
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function exportReport(config: PluginConfig, record: KintoneRecord): Promise<void> {
  const reportId = String(recordValue(record, config.outputReportIdField) || '');
  const store = String(recordValue(record, config.outputStoreField) || '');
  const currentBaseDate = String(recordValue(record, config.outputBaseDateField) || '');
  const baseDate = resolveBaseDate(config.baseDateRule, new Date(), currentBaseDate);

  if (!reportId) {
    throw new Error(
      [
        '帳票種類が未入力です。',
        'テンプレート管理アプリで初期Excelテンプレートを作成したい場合は、',
        'このアプリのプラグイン設定でモードを「テンプレート管理」に変更し、アプリを更新してください。'
      ].join('')
    );
  }
  if (!baseDate) {
    throw new Error('基準日が未入力です。');
  }

  const templateRecord = await findTemplateRecord(config, reportId);
  const reportName = String(recordValue(templateRecord, config.templateReportNameField) || reportId);
  const sources = resolveTemplateSources(config, templateRecord);
  const attachments = recordValue(templateRecord, config.templateAttachmentField);
  const fileKey = Array.isArray(attachments) ? attachments[0]?.fileKey : undefined;

  if (!fileKey) {
    throw new Error('完成版テンプレート添付が見つかりません。');
  }

  const sourceRows = await fetchSourceRows(sources, store, baseDate);
  const wholeRange = mergeSourceRanges(sourceRows, baseDate);
  const context: ReportContext = {
    reportId,
    reportName,
    store,
    baseDate,
    periodStart: wholeRange.periodStart,
    periodEnd: wholeRange.periodEnd,
    exportedAt: formatDate(new Date()),
    exporter: kintone.getLoginUser?.()?.name || ''
  };

  const templateBuffer = await downloadKintoneFile(fileKey);
  const outputBuffer = await fillReportTemplate(templateBuffer, context, sourceRows);
  downloadWorkbook(outputBuffer, buildFileName(context));
}

function resolveTemplateSources(config: PluginConfig, record: KintoneRecord): SourceAppConfig[] {
  const recordSourcesJson = String(recordValue(record, config.templateSourcesJsonField) || '');
  return parseTemplateSources(recordSourcesJson, config.sources);
}

async function fetchSourceRows(sources: SourceAppConfig[], store: string, baseDate: string): Promise<SourceRows[]> {
  const results: SourceRows[] = [];

  for (const source of sources) {
    if (!source.appId) {
      continue;
    }

    const range = sourceDateRange(source, baseDate);
    const query = buildSourceQuery(source, store, baseDate);
    const fieldCodes = source.fields.map((field) => field.code);
    const records = await getAllRecords(source.appId, query, fieldCodes);
    const rows = records.map((record) =>
      Object.fromEntries(source.fields.map((field) => [field.code, fieldDisplayValue(record, field)]))
    );

    results.push({
      source,
      rows,
      periodStart: range.start,
      periodEnd: range.end
    });
  }

  return results;
}

function sourceDateRange(source: SourceAppConfig, baseDate: string): { start: string; end: string } {
  const ranges = source.filters
    .filter((filter) => filter.valueFrom === 'dateRange')
    .map((filter) => calculateDateRange(filter.dateRule || 'sameDay', baseDate));
  const starts = ranges.map((range) => range.start).sort();
  const ends = ranges.map((range) => range.end).sort();
  return {
    start: starts[0] || baseDate,
    end: ends[ends.length - 1] || baseDate
  };
}

function mergeSourceRanges(sourceRows: SourceRows[], fallbackDate: string): { periodStart: string; periodEnd: string } {
  const starts = sourceRows.map((source) => source.periodStart).filter(Boolean).sort();
  const ends = sourceRows.map((source) => source.periodEnd).filter(Boolean).sort();
  return {
    periodStart: starts[0] || fallbackDate,
    periodEnd: ends[ends.length - 1] || fallbackDate
  };
}

async function runWithStatus(button: HTMLButtonElement, status: HTMLElement, action: () => Promise<void>): Promise<void> {
  button.disabled = true;
  status.textContent = '処理中...';

  try {
    await action();
    status.textContent = '完了しました。';
  } catch (error) {
    status.textContent = 'エラー';
    window.alert(errorMessage(error, '処理に失敗しました。'));
  } finally {
    button.disabled = false;
  }
}

function errorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (error && typeof error === 'object' && 'message' in error) {
    const message = String((error as { message: unknown }).message || '');
    const code = 'code' in error ? String((error as { code: unknown }).code || '') : '';
    return [message, code ? `エラーコード: ${code}` : ''].filter(Boolean).join('\n') || fallback;
  }

  return fallback;
}
