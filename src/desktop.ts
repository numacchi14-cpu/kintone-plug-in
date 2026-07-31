import { parsePluginConfig, parseTemplateSources } from './shared/config';
import { calculateDateRange, formatDate, resolveBaseDate } from './shared/dateRules';
import { buildFileName, createInitialTemplate, downloadWorkbook, fillReportTemplate, validateReportTemplate } from './shared/excel';
import {
  buildSourceQuery,
  buildInQuery,
  downloadKintoneFile,
  fieldDisplayValue,
  findTemplateRecord,
  getAllRecords,
  getSourceAppFields,
  recordValue,
  uploadKintoneFile
} from './shared/kintoneApi';
import type { SourceFieldImportResult } from './shared/kintoneApi';
import type {
  KintoneRecord,
  PluginConfig,
  ReportContext,
  SourceAppConfig,
  SourceFieldConfig,
  SourceFilterConfig,
  SourceLookupConfig,
  SourceSortConfig,
  SourceRows
} from './shared/types';

declare const kintone: any;

const pluginId = kintone.$PLUGIN_ID;
const pluginConfig = parsePluginConfig(kintone.plugin.app.getConfig(pluginId));
let sourceBlockSequence = 0;
const indexRecordPickerSelection = new Map<string, KintoneRecord>();

kintone.events.on(['app.record.create.show'], (event: any) => {
  if (pluginConfig.mode !== 'output') {
    return event;
  }

  const field = event.record[pluginConfig.outputBaseDateField];
  if (field) {
    field.value = resolveBaseDate('yesterday');
  }

  return event;
});

kintone.events.on(['app.record.create.show', 'app.record.edit.show'], (event: any) => {
  if (pluginConfig.mode === 'template') {
    renderSourceJsonBuilder(pluginConfig);
  } else {
    applyOutputPeriodDefaults(pluginConfig, event.record);
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
    await runWithStatus(button, status, async (setStatus) => {
      if (pluginConfig.mode === 'template') {
        await generateInitialTemplate(pluginConfig, event.record);
      } else {
        await exportReport(pluginConfig, event.record, setStatus);
      }
    });
  });

  toolbar.append(button);
  if (pluginConfig.mode === 'template') {
    const uploadButton = document.createElement('button');
    uploadButton.type = 'button';
    uploadButton.className = 'krp-button krp-button--secondary';
    uploadButton.textContent = '完成版テンプレート添付を更新';
    uploadButton.addEventListener('click', async () => {
      await runWithStatus(uploadButton, status, async (setStatus) => {
        await updateCompletedTemplateAttachment(pluginConfig, setStatus);
      });
    });
    toolbar.append(uploadButton);

    const validateButton = document.createElement('button');
    validateButton.type = 'button';
    validateButton.className = 'krp-button krp-button--secondary';
    validateButton.textContent = 'テンプレート検証';
    validateButton.addEventListener('click', async () => {
      await runWithStatus(validateButton, status, async (setStatus) => {
        await validateCompletedTemplate(pluginConfig, event.record, setStatus);
      });
    });
    toolbar.append(validateButton);
  } else {
    const checkButton = document.createElement('button');
    checkButton.type = 'button';
    checkButton.className = 'krp-button krp-button--secondary';
    checkButton.textContent = '出力前チェック';
    checkButton.addEventListener('click', async () => {
      await runWithStatus(checkButton, status, async (setStatus) => {
        await validateOutputReadiness(pluginConfig, event.record, setStatus);
      });
    });
    toolbar.append(checkButton);
  }
  toolbar.append(status);
  header.appendChild(toolbar);
  return event;
});

kintone.events.on(['app.record.index.show'], (event: any) => {
  if (pluginConfig.mode !== 'output') {
    return event;
  }

  renderIndexRecordPicker(pluginConfig, (event.records ?? []) as KintoneRecord[]);

  const header = kintone.app.getHeaderMenuSpaceElement();
  if (!header || header.querySelector('[data-krp-index-toolbar="true"]')) {
    return event;
  }

  const toolbar = document.createElement('span');
  toolbar.className = 'krp-toolbar';
  toolbar.dataset.krpIndexToolbar = 'true';

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'krp-button';
  button.textContent = '選択レコードをExcel出力';
  button.title = '下に表示されるチェック欄でレコードを選択してからクリックしてください';

  const status = document.createElement('span');
  status.className = 'krp-status';
  status.textContent = '帳票出力モード';

  button.addEventListener('click', async () => {
    const records = Array.from(indexRecordPickerSelection.values());
    if (records.length > 1 && !window.confirm(`選択した${records.length}件をExcel出力します。よろしいですか？`)) {
      return;
    }
    await runWithStatus(button, status, async (setStatus) => {
      await exportSelectedReports(pluginConfig, records, setStatus);
    });
  });

  toolbar.append(button, status);
  header.appendChild(toolbar);
  return event;
});

function renderIndexRecordPicker(config: PluginConfig, records: KintoneRecord[]): void {
  const space = kintone.app.getHeaderSpaceElement?.();
  if (!space) {
    return;
  }

  indexRecordPickerSelection.clear();
  space.querySelector('[data-krp-record-picker="true"]')?.remove();

  if (!records.length) {
    return;
  }

  const panel = document.createElement('div');
  panel.className = 'krp-record-picker';
  panel.dataset.krpRecordPicker = 'true';

  const controls = document.createElement('div');
  controls.className = 'krp-record-picker__controls';

  const hint = document.createElement('span');
  hint.className = 'krp-record-picker__hint';
  hint.textContent = 'Excel出力するレコードにチェック:';

  const selectAllButton = document.createElement('button');
  selectAllButton.type = 'button';
  selectAllButton.textContent = '全選択';

  const clearAllButton = document.createElement('button');
  clearAllButton.type = 'button';
  clearAllButton.textContent = '全解除';

  controls.append(hint, selectAllButton, clearAllButton);

  const list = document.createElement('div');
  list.className = 'krp-record-picker__list';

  const checkboxes: HTMLInputElement[] = [];
  for (const record of records) {
    const id = String(recordValue(record, '$id') || '');
    if (!id) {
      continue;
    }

    const label = document.createElement('label');
    label.className = 'krp-record-picker__item';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) {
        indexRecordPickerSelection.set(id, record);
      } else {
        indexRecordPickerSelection.delete(id);
      }
    });
    checkboxes.push(checkbox);

    const reportType = String(recordValue(record, config.outputReportIdField) || '');
    const store = String(recordValue(record, config.outputStoreField) || '');
    const baseDate = resolveBaseDate('yesterday');

    const text = document.createElement('span');
    text.textContent = `No.${id} ${[reportType, store, `基準日:${baseDate}`].filter(Boolean).join(' / ')}`;

    label.append(checkbox, text);
    list.append(label);
  }

  selectAllButton.addEventListener('click', () => {
    checkboxes.forEach((checkbox) => {
      checkbox.checked = true;
      checkbox.dispatchEvent(new Event('change'));
    });
  });
  clearAllButton.addEventListener('click', () => {
    checkboxes.forEach((checkbox) => {
      checkbox.checked = false;
      checkbox.dispatchEvent(new Event('change'));
    });
  });

  panel.append(controls, list);
  space.appendChild(panel);
}

async function exportSelectedReports(
  config: PluginConfig,
  records: KintoneRecord[],
  setStatus: (message: string) => void = () => undefined
): Promise<void> {
  if (!records.length) {
    throw new Error('レコードが選択されていません。一覧上部のチェック欄でレコードを選択してください。');
  }

  const failures: string[] = [];
  for (const [index, record] of records.entries()) {
    const progress = `${index + 1}/${records.length}件目`;
    try {
      await exportReport(config, record, (message) => setStatus(`${progress}: ${message}`));
    } catch (error) {
      failures.push(`${progress}: ${errorMessage(error, '出力に失敗しました。')}`);
    }
    if (index < records.length - 1) {
      await wait(300);
    }
  }

  if (failures.length) {
    throw new Error([`${records.length}件中${failures.length}件の出力に失敗しました。`, ...failures].join('\n'));
  }
}

function wait(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function generateInitialTemplate(config: PluginConfig, record: KintoneRecord): Promise<void> {
  const reportName = String(recordValue(record, config.templateReportNameField) || 'Excel帳票テンプレート');
  const sources = resolveTemplateSources(config, record);
  const buffer = await createInitialTemplate({ ...config, sources }, reportName);
  downloadWorkbook(buffer, `${reportName.replace(/[\\/:*?"<>|]/g, '_')}_初期テンプレート.xlsx`);
}

async function updateCompletedTemplateAttachment(
  config: PluginConfig,
  setStatus: (message: string) => void = () => undefined
): Promise<void> {
  const file = await chooseXlsxFile();
  if (!file) {
    setStatus('キャンセルしました。');
    return;
  }
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    throw new Error('完成版テンプレートはxlsx形式のファイルを選択してください。');
  }

  const appId = kintone.app.getId?.();
  const recordId = kintone.app.record.getId?.();
  if (!appId || !recordId) {
    throw new Error('レコードIDを取得できません。レコード詳細画面で実行してください。');
  }

  setStatus('完成版テンプレートをアップロード中...');
  const fileKey = await uploadKintoneFile(file);
  const record: Record<string, { value: Array<{ fileKey: string }> }> = {
    [config.templateAttachmentField]: { value: [{ fileKey }] }
  };

  setStatus('添付フィールドを更新中...');
  await kintone.api(kintone.api.url('/k/v1/record.json', true), 'PUT', {
    app: appId,
    id: recordId,
    record
  });

  setStatus('完成版テンプレート添付を更新しました。');
  window.alert('完成版テンプレート添付を更新しました。画面を再読み込みすると添付欄にも反映されます。');
}

function chooseXlsxFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    let resolved = false;
    const finish = (file: File | null) => {
      if (resolved) {
        return;
      }
      resolved = true;
      window.removeEventListener('focus', handleFocus);
      input.remove();
      resolve(file);
    };
    const handleFocus = () => {
      window.setTimeout(() => {
        if (!input.files?.length) {
          finish(null);
        }
      }, 300);
    };

    input.type = 'file';
    input.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    input.style.display = 'none';
    input.addEventListener(
      'change',
      () => {
        finish(input.files?.[0] ?? null);
      },
      { once: true }
    );
    document.body.appendChild(input);
    window.addEventListener('focus', handleFocus);
    input.click();
  });
}

async function validateCompletedTemplate(
  config: PluginConfig,
  fallbackRecord: KintoneRecord,
  setStatus: (message: string) => void = () => undefined
): Promise<void> {
  setStatus('テンプレート管理レコードを確認中...');
  const record = (await fetchCurrentRecord()).record || fallbackRecord;
  const sources = resolveTemplateSources(config, record);
  const attachments = recordValue(record, config.templateAttachmentField);
  const fileKey = resolveCompletedTemplateFileKey(attachments);

  setStatus('取得元アプリ設定を検証中...');
  await validateSourceConfigs(config, sources);

  setStatus('完成版テンプレートを取得中...');
  const templateBuffer = await downloadKintoneFile(fileKey);

  setStatus('Excelテンプレートを検証中...');
  await validateReportTemplate(templateBuffer, sources);

  setStatus('テンプレート検証が完了しました。');
  window.alert('テンプレート検証が完了しました。出力に必要なシートと列見出しは揃っています。');
}

async function fetchCurrentRecord(): Promise<{ record?: KintoneRecord }> {
  const appId = kintone.app.getId?.();
  const recordId = kintone.app.record.getId?.();
  if (!appId || !recordId) {
    return {};
  }

  return kintone.api(kintone.api.url('/k/v1/record.json', true), 'GET', {
    app: appId,
    id: recordId
  });
}

function renderBaseDateUpdateButton(config: PluginConfig): void {
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
  button.textContent = '基準日を昨日に更新';
  button.title = '基準日を実行日の前日に更新します';
  button.addEventListener('click', () => {
    const current = kintone.app.record.get();
    const field = current?.record?.[config.outputBaseDateField];
    if (!field) {
      window.alert(`基準日フィールド「${config.outputBaseDateField}」がフォームにありません。`);
      return;
    }

    field.value = resolveBaseDate('yesterday', new Date(), '');
    applyOutputPeriodDefaults(config, current.record, true);
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
      <button type="button" data-krp-action="import-output-fields">出力アプリのフィールド一覧を取得</button>
      <button type="button" data-krp-action="apply-json">JSONをフィールドへ反映</button>
      <button type="button" data-krp-action="load-json">現在のJSONを読み込む</button>
      <span class="krp-builder__source-status" data-krp-builder="outputFieldStatus"></span>
    </div>
    <datalist id="krp-output-field-codes" data-krp-builder="outputFieldCodes"></datalist>
    <div data-krp-builder="sources"></div>
    <p class="krp-builder__note">反映後は、このレコードを保存してください。取得元アプリは複数追加できます。列順は各ブロック内の出力フィールドの行順です。出力アプリの任意フィールドを抽出条件に使う場合は、プラグイン設定で紐づく帳票出力アプリIDを設定してください。</p>
  `;

  const addSourceButton = panel.querySelector('[data-krp-action="add-source"]');
  const importOutputFieldsButton = panel.querySelector('[data-krp-action="import-output-fields"]');
  const applyButton = panel.querySelector('[data-krp-action="apply-json"]');
  const loadButton = panel.querySelector('[data-krp-action="load-json"]');

  addSourceButton?.addEventListener('click', () => addBuilderSourceBlock(panel));
  importOutputFieldsButton?.addEventListener('click', async () => importOutputFields(config, panel));
  applyButton?.addEventListener('click', () => applyBuilderJson(config, panel));
  loadButton?.addEventListener('click', () => loadBuilderJson(config, panel));

  header.appendChild(panel);
  addBuilderSourceBlock(panel);
}

async function importOutputFields(config: PluginConfig, panel: HTMLElement): Promise<void> {
  const status = panel.querySelector('[data-krp-builder="outputFieldStatus"]');
  const datalist = panel.querySelector('[data-krp-builder="outputFieldCodes"]');
  if (!(datalist instanceof HTMLDataListElement)) {
    return;
  }

  if (!config.outputAppId) {
    window.alert('プラグイン設定で「紐づく帳票出力アプリID」を入力してください。');
    return;
  }

  try {
    if (status) {
      status.textContent = '出力アプリのフィールド一覧を取得中...';
    }
    const result = await getSourceAppFields(config.outputAppId);
    datalist.replaceChildren(
      ...result.fields.map((field) => {
        const option = document.createElement('option');
        option.value = field.code;
        option.label = field.label;
        return option;
      })
    );
    if (status) {
      status.textContent = `出力アプリ ${config.outputAppId}: ${result.fields.length}件取得 / ${result.skippedCount}件除外`;
    }
  } catch (error) {
    if (status) {
      status.textContent = '';
    }
    window.alert(error instanceof Error ? error.message : '出力アプリのフィールド一覧を取得できません。');
  }
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
      <label>Excelテーブル名<input data-krp-source="tableName" placeholder="例: tbl_actual" value="${escapeHtml(source?.tableName ?? '')}"></label>
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
    <div class="krp-builder__rules">
      <div class="krp-builder__rules-head">
        <strong>マスタ参照</strong>
        <button type="button" data-krp-action="add-lookup">参照を追加</button>
      </div>
      <div data-krp-source="lookups"></div>
      <p class="krp-builder__note">店舗IDなどのキーでマスタアプリを参照し、店舗名や営業部などを補完します。補完フィールドはJSON配列で指定します。</p>
    </div>
    <div class="krp-builder__fields">
      <div class="krp-builder__fields-head">
        <strong>出力フィールド</strong>
        <button type="button" data-krp-action="add-field">行を追加</button>
        <button type="button" data-krp-action="select-all-fields">削除対象を全選択</button>
        <button type="button" data-krp-action="clear-all-fields">削除対象を全解除</button>
        <button type="button" data-krp-action="remove-selected-fields">選択行を削除へ移動</button>
        <button type="button" data-krp-action="restore-all-fields">削除フィールドをすべて戻す</button>
      </div>
      <div class="krp-builder__field-panes">
        <section class="krp-builder__field-pane">
          <div class="krp-builder__field-pane-head">
            <strong>出力フィールド</strong>
            <span data-krp-field-count="active"></span>
          </div>
          <div data-krp-source="fields"></div>
        </section>
        <section class="krp-builder__field-pane">
          <div class="krp-builder__field-pane-head">
            <strong>削除フィールド</strong>
            <span data-krp-field-count="inactive"></span>
          </div>
          <div data-krp-source="inactiveFields"></div>
        </section>
      </div>
      <datalist id="${sourceBlockId}" data-krp-source="fieldCodes"></datalist>
      <p class="krp-builder__note">出力フィールドだけJSONへ反映されます。削除フィールドは後から戻せます。一度すべて削除して、戻す順番で列順を整理できます。</p>
    </div>
  `;

  sourceBlock.querySelector('[data-krp-action="import-fields"]')?.addEventListener('click', async () => {
    await importBuilderFields(sourceBlock);
  });
  sourceBlock.querySelector('[data-krp-action="add-filter"]')?.addEventListener('click', () => addBuilderFilterRow(sourceBlock));
  sourceBlock.querySelector('[data-krp-action="add-sort"]')?.addEventListener('click', () => addBuilderSortRow(sourceBlock));
  sourceBlock.querySelector('[data-krp-action="add-lookup"]')?.addEventListener('click', () => addBuilderLookupRow(sourceBlock));
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
  sourceBlock.querySelector('[data-krp-action="restore-all-fields"]')?.addEventListener('click', () => {
    restoreAllBuilderFields(sourceBlock);
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
  source?.lookups.forEach((lookup) => addBuilderLookupRow(sourceBlock, lookup));

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
      <option value="outputField">出力アプリの任意フィールド</option>
      <option value="fixed">固定値</option>
    </select>
    <select data-krp-filter="dateRule">
      <optgroup label="日">
        <option value="sameDay">基準日</option>
        <option value="previousDay">基準日の前日</option>
        <option value="nextDay">基準日の翌日</option>
      </optgroup>
      <optgroup label="週（日曜から土曜）">
        <option value="baseWeek">基準日の週</option>
        <option value="previousWeek">基準日の前週</option>
        <option value="nextWeek">基準日の翌週</option>
      </optgroup>
      <optgroup label="月">
        <option value="baseMonthStart">基準日の同月1日</option>
        <option value="baseMonthEnd">基準日の同月末</option>
        <option value="monthStartToBaseDate">基準日の同月1日から基準日</option>
        <option value="baseMonth">基準日の同月</option>
        <option value="previousMonthStart">基準日の前月1日</option>
        <option value="previousMonthEnd">基準日の前月末</option>
        <option value="previousMonth">基準日の前月</option>
        <option value="nextMonthStart">基準日の翌月1日</option>
        <option value="nextMonthEnd">基準日の翌月末</option>
        <option value="nextMonth">基準日の翌月</option>
        <option value="baseMonthToNextMonthEnd">基準日の同月1日から翌月末</option>
      </optgroup>
      <optgroup label="年">
        <option value="baseYearStart">基準日の同年1月1日</option>
        <option value="baseYearEnd">基準日の同年12月31日</option>
        <option value="yearStartToBaseDate">基準日の同年1月1日から基準日</option>
        <option value="baseYear">基準日の同年</option>
        <option value="previousYearStart">基準日の前年1月1日</option>
        <option value="previousYearEnd">基準日の前年12月31日</option>
        <option value="previousYear">基準日の前年</option>
        <option value="nextYearStart">基準日の翌年1月1日</option>
        <option value="nextYearEnd">基準日の翌年12月31日</option>
        <option value="nextYear">基準日の翌年</option>
      </optgroup>
      <optgroup label="前年比較">
        <option value="sameDayPreviousYear">基準日の前年同日</option>
        <option value="sameMonthPreviousYear">基準日の前年同月</option>
        <option value="previousMonthPreviousYear">基準日の前月の前年同月</option>
      </optgroup>
    </select>
    <input data-krp-filter="outputField" list="krp-output-field-codes" placeholder="出力アプリのフィールドコード" value="${escapeHtml(filter.outputField ?? '')}">
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
  const outputField = row.querySelector('[data-krp-filter="outputField"]');
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
  if (outputField instanceof HTMLInputElement) {
    outputField.value = filter.outputField ?? '';
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
  const outputField = row.querySelector('[data-krp-filter="outputField"]');
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

  if (outputField instanceof HTMLInputElement) {
    const usesOutputField = valueFrom.value === 'outputField';
    outputField.hidden = !usesOutputField;
    outputField.disabled = !usesOutputField;
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
    <button type="button" class="krp-builder__drag-handle" data-krp-drag-handle aria-label="ドラッグで並び替え" title="ドラッグで並び替え">↕</button>
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
  enableBuilderRowDrag(row, sorts, 'krp-builder__sort-row');
  sorts.appendChild(row);
}

function enableBuilderRowDrag(row: HTMLElement, container: Element, rowClassName: string): void {
  const handle = row.querySelector('[data-krp-drag-handle]');
  if (!(handle instanceof HTMLElement)) {
    return;
  }

  handle.draggable = true;
  handle.addEventListener('dragstart', (event) => {
    row.classList.add('krp-builder__row--dragging');
    event.dataTransfer?.setData('text/plain', '');
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
    }
  });
  handle.addEventListener('dragend', () => {
    row.classList.remove('krp-builder__row--dragging');
    container.querySelectorAll('.krp-builder__row--drop-before').forEach((element) => {
      element.classList.remove('krp-builder__row--drop-before');
    });
  });

  row.addEventListener('dragover', (event) => {
    const draggingRow = container.querySelector('.krp-builder__row--dragging');
    if (!(draggingRow instanceof HTMLElement) || draggingRow === row) {
      return;
    }
    if (!draggingRow.classList.contains(rowClassName)) {
      return;
    }

    event.preventDefault();
    const insertBefore = event.clientY < row.getBoundingClientRect().top + row.getBoundingClientRect().height / 2;
    row.classList.toggle('krp-builder__row--drop-before', insertBefore);
    if (insertBefore) {
      row.before(draggingRow);
    } else {
      row.after(draggingRow);
    }
  });

  row.addEventListener('dragleave', () => {
    row.classList.remove('krp-builder__row--drop-before');
  });
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

function addBuilderLookupRow(sourceBlock: HTMLElement, lookup: Partial<SourceLookupConfig> = {}): void {
  const lookups = sourceBlock.querySelector('[data-krp-source="lookups"]');
  if (!lookups) {
    return;
  }

  const row = document.createElement('div');
  row.className = 'krp-builder__lookup-row';
  row.innerHTML = `
    <input data-krp-lookup="sourceField" list="${fieldCodeListId(sourceBlock)}" placeholder="取得元キー" value="${escapeHtml(lookup.sourceField ?? '')}">
    <input data-krp-lookup="masterAppId" inputmode="numeric" placeholder="マスタアプリID" value="${escapeHtml(lookup.masterAppId ?? '')}">
    <input data-krp-lookup="masterKeyField" placeholder="マスタキー" value="${escapeHtml(lookup.masterKeyField ?? '')}">
    <textarea data-krp-lookup="masterFields" spellcheck="false" placeholder='[{"code":"店舗名","label":"店舗名"}]'>${escapeHtml(JSON.stringify(lookup.masterFields ?? [], null, 2))}</textarea>
    <button type="button" data-krp-action="remove-lookup">削除</button>
  `;

  row.querySelector('[data-krp-action="remove-lookup"]')?.addEventListener('click', () => row.remove());
  lookups.appendChild(row);
}

function fieldCodeListId(sourceBlock: Element): string {
  return sourceBlock.querySelector('[data-krp-source="fieldCodes"]')?.id || '';
}

function addBuilderFieldRow(
  sourceBlock: HTMLElement,
  field: { code?: string; label?: string; type?: string } = {},
  selected = true,
  target: 'active' | 'inactive' = 'active'
): void {
  const fields = target === 'active' ? activeFieldsContainer(sourceBlock) : inactiveFieldsContainer(sourceBlock);
  if (!fields) {
    return;
  }

  const row = document.createElement('div');
  row.className = `krp-builder__field-row${target === 'inactive' ? ' krp-builder__field-row--inactive' : ''}`;
  row.dataset.krpFieldState = target;
  row.innerHTML = fieldRowHtml(field, selected, target);

  bindBuilderFieldRow(sourceBlock, row, target);
  fields.appendChild(row);
  refreshFieldCodeSuggestions(sourceBlock);
  refreshBuilderFieldCounts(sourceBlock);
}

function fieldRowHtml(
  field: { code?: string; label?: string; type?: string },
  selected: boolean,
  target: 'active' | 'inactive'
): string {
  const active = target === 'active';
  return `
    ${
      active
        ? '<button type="button" class="krp-builder__drag-handle" data-krp-drag-handle aria-label="ドラッグで列順を変更" title="ドラッグで列順を変更">↕</button>'
        : '<span class="krp-builder__field-placeholder"></span>'
    }
    ${
      active
        ? `<input type="checkbox" data-krp-field="selected" aria-label="削除対象" title="選択行を削除フィールドへ移動"${selected ? ' checked' : ''}>`
        : '<span class="krp-builder__field-placeholder"></span>'
    }
    <input data-krp-field="code" placeholder="フィールドコード" value="${escapeHtml(field.code ?? '')}">
    <input data-krp-field="label" placeholder="Excel見出し" value="${escapeHtml(field.label ?? '')}">
    <select data-krp-field="type">
      <option value="text"${field.type === 'text' || !field.type ? ' selected' : ''}>文字列</option>
      <option value="number"${field.type === 'number' ? ' selected' : ''}>数値</option>
      <option value="date"${field.type === 'date' ? ' selected' : ''}>日付</option>
      <option value="datetime"${field.type === 'datetime' ? ' selected' : ''}>日時</option>
      <option value="boolean"${field.type === 'boolean' ? ' selected' : ''}>真偽値</option>
    </select>
    <div class="krp-builder__field-actions">
      ${
        active
          ? '<button type="button" data-krp-action="move-field-up" aria-label="上へ移動" title="上へ移動">↑</button><button type="button" data-krp-action="move-field-down" aria-label="下へ移動" title="下へ移動">↓</button><button type="button" data-krp-action="remove-field">削除</button>'
          : '<button type="button" data-krp-action="restore-field">戻す</button><button type="button" data-krp-action="remove-field-permanently">完全削除</button>'
      }
    </div>
  `;
}

function bindBuilderFieldRow(sourceBlock: HTMLElement, row: HTMLElement, target: 'active' | 'inactive'): void {
  row.querySelector('[data-krp-field="code"]')?.addEventListener('input', () => refreshFieldCodeSuggestions(sourceBlock));
  row.querySelector('[data-krp-action="move-field-up"]')?.addEventListener('click', () => {
    row.previousElementSibling?.before(row);
  });
  row.querySelector('[data-krp-action="move-field-down"]')?.addEventListener('click', () => {
    row.nextElementSibling?.after(row);
  });
  row.querySelector('[data-krp-action="remove-field"]')?.addEventListener('click', () => {
    moveBuilderFieldRow(sourceBlock, row, 'inactive');
  });
  row.querySelector('[data-krp-action="restore-field"]')?.addEventListener('click', () => {
    moveBuilderFieldRow(sourceBlock, row, 'active');
  });
  row.querySelector('[data-krp-action="remove-field-permanently"]')?.addEventListener('click', () => {
    row.remove();
    refreshFieldCodeSuggestions(sourceBlock);
    refreshBuilderFieldCounts(sourceBlock);
  });

  const fields = activeFieldsContainer(sourceBlock);
  if (target === 'active' && fields) {
    enableBuilderRowDrag(row, fields, 'krp-builder__field-row');
  }
}

function moveBuilderFieldRow(sourceBlock: HTMLElement, row: Element, target: 'active' | 'inactive'): void {
  if (!(row instanceof HTMLElement)) {
    return;
  }

  const field = readBuilderFieldRow(row);
  row.remove();
  addBuilderFieldRow(sourceBlock, field, false, target);
  refreshFieldCodeSuggestions(sourceBlock);
}

function activeFieldsContainer(sourceBlock: Element): Element | null {
  return sourceBlock.querySelector('[data-krp-source="fields"]');
}

function inactiveFieldsContainer(sourceBlock: Element): Element | null {
  return sourceBlock.querySelector('[data-krp-source="inactiveFields"]');
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
    const currentActiveFields = readBuilderFieldRows(sourceBlock, false);
    const currentInactiveFields = Array.from(
      sourceBlock.querySelectorAll('[data-krp-source="inactiveFields"] .krp-builder__field-row')
    ).map((row) => readBuilderFieldRow(row));
    const hadExistingFields = currentActiveFields.length > 0 || currentInactiveFields.length > 0;

    const result = await getSourceAppFields(appId);

    if (!result.fields.length) {
      throw new Error('出力に対応したフィールドが見つかりません。');
    }

    const liveFieldsByCode = new Map(result.fields.map((field) => [field.code, field]));

    const fieldsContainer = sourceBlock.querySelector('[data-krp-source="fields"]');
    const inactiveFieldsContainerElement = sourceBlock.querySelector('[data-krp-source="inactiveFields"]');
    if (!fieldsContainer || !inactiveFieldsContainerElement) {
      return;
    }
    fieldsContainer.innerHTML = '';
    inactiveFieldsContainerElement.innerHTML = '';

    const removedLabels: string[] = [];

    if (hadExistingFields) {
      // 既存の設定(並び順・ラベル・型)はそのまま維持し、kintone側になくなったフィールドだけ除外する。
      const knownCodes = new Set<string>();
      currentActiveFields.forEach((field) => {
        if (!liveFieldsByCode.has(field.code)) {
          removedLabels.push(field.label || field.code);
          return;
        }
        knownCodes.add(field.code);
        addBuilderFieldRow(sourceBlock, { code: field.code, label: field.label, type: field.type }, field.selected, 'active');
      });
      currentInactiveFields.forEach((field) => {
        if (!liveFieldsByCode.has(field.code)) {
          removedLabels.push(field.label || field.code);
          return;
        }
        knownCodes.add(field.code);
        addBuilderFieldRow(sourceBlock, { code: field.code, label: field.label, type: field.type }, field.selected, 'inactive');
      });

      // kintone側の新規フィールドは、既存の出力フィールド構成を崩さないよう削除フィールド側に追加する。
      result.fields.forEach((field) => {
        if (knownCodes.has(field.code)) {
          return;
        }
        addBuilderFieldRow(sourceBlock, field, true, 'inactive');
      });
    } else {
      result.fields.forEach((field) => {
        addBuilderFieldRow(sourceBlock, field, true, 'active');
      });
    }

    refreshFieldCodeSuggestions(sourceBlock);
    if (status) {
      const parts = [`${result.fields.length}件を取得`];
      if (result.skippedCount) {
        parts.push(`非対応${result.skippedCount}件を除外`);
      }
      if (removedLabels.length) {
        parts.push(`kintone側になくなった${removedLabels.length}件を除外`);
      }
      status.textContent = parts.join('、');
    }
    if (removedLabels.length) {
      window.alert(
        `kintone側に存在しなくなったため、次のフィールドを設定から削除しました。\n${removedLabels.join('\n')}`
      );
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
  sourceBlock: Element,
  includeInactive = false
): Array<SourceFieldConfig & { selected: boolean; state: 'active' | 'inactive' }> {
  const selector = includeInactive
    ? '.krp-builder__field-row'
    : '[data-krp-source="fields"] .krp-builder__field-row';
  return Array.from(sourceBlock.querySelectorAll(selector))
    .map((row) => readBuilderFieldRow(row))
    .filter((field) => field.code);
}

function readBuilderFieldRow(row: Element): SourceFieldConfig & { selected: boolean; state: 'active' | 'inactive' } {
  const code = getRowValue(row, 'code');
  const label = getRowValue(row, 'label') || code;
  const type = getRowValue(row, 'type') as SourceFieldConfig['type'];
  const selectedElement = row.querySelector('[data-krp-field="selected"]');
  const state = row instanceof HTMLElement && row.dataset.krpFieldState === 'inactive' ? 'inactive' : 'active';
  return {
    code,
    label,
    type,
    selected: selectedElement instanceof HTMLInputElement ? selectedElement.checked : true,
    state
  };
}

function setBuilderFieldSelection(sourceBlock: Element, selected: boolean): void {
  sourceBlock.querySelectorAll('[data-krp-source="fields"] [data-krp-field="selected"]').forEach((element) => {
    if (element instanceof HTMLInputElement) {
      element.checked = selected;
    }
  });
}

function removeSelectedBuilderFields(sourceBlock: HTMLElement): void {
  sourceBlock.querySelectorAll('[data-krp-source="fields"] .krp-builder__field-row').forEach((row) => {
    const checkbox = row.querySelector('[data-krp-field="selected"]');
    if (checkbox instanceof HTMLInputElement && checkbox.checked) {
      moveBuilderFieldRow(sourceBlock, row, 'inactive');
    }
  });
  refreshFieldCodeSuggestions(sourceBlock);
}

function restoreAllBuilderFields(sourceBlock: HTMLElement): void {
  Array.from(sourceBlock.querySelectorAll('[data-krp-source="inactiveFields"] .krp-builder__field-row')).forEach((row) => {
    moveBuilderFieldRow(sourceBlock, row, 'active');
  });
}

function refreshFieldCodeSuggestions(sourceBlock: Element): void {
  const datalist = sourceBlock.querySelector('[data-krp-source="fieldCodes"]');
  if (!(datalist instanceof HTMLDataListElement)) {
    return;
  }

  const codes = readBuilderFieldRows(sourceBlock, true).map((field) => field.code);
  datalist.replaceChildren(
    ...codes.map((code) => {
      const option = document.createElement('option');
      option.value = code;
      return option;
    })
  );
  refreshBuilderFieldCounts(sourceBlock);
}

function refreshBuilderFieldCounts(sourceBlock: Element): void {
  const activeCount = sourceBlock.querySelector('[data-krp-field-count="active"]');
  const inactiveCount = sourceBlock.querySelector('[data-krp-field-count="inactive"]');
  if (activeCount) {
    activeCount.textContent = `${sourceBlock.querySelectorAll('[data-krp-source="fields"] .krp-builder__field-row').length}件`;
  }
  if (inactiveCount) {
    inactiveCount.textContent = `${sourceBlock.querySelectorAll('[data-krp-source="inactiveFields"] .krp-builder__field-row').length}件`;
  }
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
  const lookups: SourceLookupConfig[] = [];

  Array.from(sourceBlock.querySelectorAll('[data-krp-source="fields"] .krp-builder__field-row')).forEach((row) => {
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
    if (valueFrom === 'outputField') {
      filter.outputField = getBuilderControlValue(row, 'filter', 'outputField');
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

  sourceBlock.querySelectorAll('.krp-builder__lookup-row').forEach((row) => {
    const sourceField = getBuilderControlValue(row, 'lookup', 'sourceField');
    const masterAppId = getBuilderControlValue(row, 'lookup', 'masterAppId');
    const masterKeyField = getBuilderControlValue(row, 'lookup', 'masterKeyField');
    if (!sourceField || !masterAppId || !masterKeyField) {
      return;
    }

    lookups.push({
      sourceField,
      masterAppId,
      masterKeyField,
      masterFields: parseBuilderMasterFields(getBuilderControlValue(row, 'lookup', 'masterFields'))
    });
  });

  return {
    key: toSourceKey(label || getSourceValue(sourceBlock, 'sheetName') || `source_${sourceNumber}`),
    label,
    appId: getSourceValue(sourceBlock, 'appId'),
    sheetName: getSourceValue(sourceBlock, 'sheetName'),
    tableName: getSourceValue(sourceBlock, 'tableName') || undefined,
    fields,
    filters,
    sorts,
    lookups
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

function getBuilderControlValue(row: Element, group: 'filter' | 'sort' | 'lookup', name: string): string {
  const element = row.querySelector(`[data-krp-${group}="${name}"]`);
  return element instanceof HTMLInputElement || element instanceof HTMLSelectElement || element instanceof HTMLTextAreaElement
    ? element.value.trim()
    : '';
}

function parseBuilderMasterFields(value: string): SourceFieldConfig[] {
  if (!value) {
    return [];
  }

  try {
    const parsed = JSON.parse(value) as SourceFieldConfig[];
    return Array.isArray(parsed)
      ? parsed
          .filter((field) => field?.code)
          .map((field) => {
            const type = ['text', 'number', 'date', 'datetime', 'boolean'].includes(String(field.type))
              ? field.type
              : undefined;
            return {
              code: String(field.code),
              label: String(field.label || field.code),
              ...(type ? { type } : {})
            };
          })
      : [];
  } catch {
    window.alert('マスタ参照の補完フィールドJSONが不正です。');
    return [];
  }
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

async function exportReport(
  config: PluginConfig,
  record: KintoneRecord,
  setStatus: (message: string) => void = () => undefined
): Promise<void> {
  setStatus('出力条件を確認中...');
  const { reportId, store, baseDate } = resolveOutputConditions(config, record);

  setStatus('テンプレート設定を取得中...');
  const templateRecord = await findTemplateRecord(config, reportId);
  const reportName = String(recordValue(templateRecord, config.templateReportNameField) || reportId);
  const sources = resolveTemplateSources(config, templateRecord);
  const attachments = recordValue(templateRecord, config.templateAttachmentField);
  const fileKey = resolveCompletedTemplateFileKey(attachments);

  setStatus('取得元アプリ設定を検証中...');
  await validateSourceConfigs(config, sources);

  const sourceRows = await fetchSourceRows(sources, store, baseDate, record, setStatus);
  const wholeRange = mergeSourceRanges(sourceRows, baseDate);
  const exportedAt = formatDateTime(new Date());
  const context: ReportContext = {
    reportId,
    reportName,
    store,
    baseDate,
    periodStart: wholeRange.periodStart,
    periodEnd: wholeRange.periodEnd,
    exportedAt,
    exporter: kintone.getLoginUser?.()?.name || ''
  };

  setStatus('Excelテンプレートを取得中...');
  const templateBuffer = await downloadKintoneFile(fileKey);
  setStatus('Excel帳票を作成中...');
  const outputBuffer = await fillReportTemplate(templateBuffer, context, sourceRows);
  const fileName = buildFileName(context);
  setStatus('Excelをダウンロード中...');
  downloadWorkbook(outputBuffer, fileName);
  setStatus('出力履歴を保存中...');
  await saveOutputHistory(config, context, fileName);
}

async function validateOutputReadiness(
  config: PluginConfig,
  record: KintoneRecord,
  setStatus: (message: string) => void = () => undefined
): Promise<void> {
  setStatus('出力条件を確認中...');
  const { reportId, store, baseDate } = resolveOutputConditions(config, record);

  setStatus('テンプレート設定を取得中...');
  const templateRecord = await findTemplateRecord(config, reportId);
  const reportName = String(recordValue(templateRecord, config.templateReportNameField) || reportId);
  const sources = resolveTemplateSources(config, templateRecord);
  const attachments = recordValue(templateRecord, config.templateAttachmentField);
  const fileKey = resolveCompletedTemplateFileKey(attachments);

  setStatus('取得元アプリ設定を検証中...');
  await validateSourceConfigs(config, sources);

  setStatus('取得クエリを確認中...');
  sources.forEach((source) => {
    sourceDateRange(source, baseDate);
    buildSourceQuery(source, store, baseDate, record);
  });

  setStatus('Excelテンプレートを取得中...');
  const templateBuffer = await downloadKintoneFile(fileKey);

  setStatus('Excelテンプレートを検証中...');
  await validateReportTemplate(templateBuffer, sources);

  setStatus('出力前チェックが完了しました。');
  window.alert(`出力前チェックが完了しました。\n帳票: ${reportName}\n取得元アプリ: ${sources.length}件`);
}

function resolveOutputConditions(
  config: PluginConfig,
  record: KintoneRecord
): { reportId: string; store: string; baseDate: string } {
  const reportId = String(recordValue(record, config.outputReportIdField) || '');
  const store = String(recordValue(record, config.outputStoreField) || '');
  const baseDate = resolveBaseDate('yesterday', new Date(), '');

  if (!reportId) {
    throw new Error(
      [
        '帳票種類が未入力です。',
        'テンプレート管理アプリで初期Excelテンプレートを作成したい場合は、',
        'このアプリのプラグイン設定でモードを「テンプレート管理」に変更し、アプリを更新してください。'
      ].join('')
    );
  }
  return { reportId, store, baseDate };
}

function resolveCompletedTemplateFileKey(attachments: unknown): string {
  if (!Array.isArray(attachments) || attachments.length === 0) {
    throw new Error('完成版テンプレート添付が見つかりません。');
  }

  if (attachments.length > 1) {
    throw new Error('完成版テンプレート添付が複数あります。1つだけ添付してください。');
  }

  const fileKey = attachments[0]?.fileKey;
  if (!fileKey) {
    throw new Error('完成版テンプレート添付のfileKeyを取得できません。');
  }

  return String(fileKey);
}

function resolveTemplateSources(config: PluginConfig, record: KintoneRecord): SourceAppConfig[] {
  const recordSourcesJson = String(recordValue(record, config.templateSourcesJsonField) || '');
  return parseTemplateSources(recordSourcesJson, config.sources);
}

async function fetchSourceRows(
  sources: SourceAppConfig[],
  store: string,
  baseDate: string,
  outputRecord: KintoneRecord,
  setStatus: (message: string) => void = () => undefined
): Promise<SourceRows[]> {
  const results: SourceRows[] = [];

  for (const [index, source] of sources.entries()) {
    if (!source.appId) {
      continue;
    }

    const sourceLabel = source.label || source.key || `取得元${index + 1}`;
    setStatus(`取得中: ${sourceLabel}`);
    const range = sourceDateRange(source, baseDate);
    const query = buildSourceQuery(source, store, baseDate, outputRecord);
    const lookupFieldCodes = new Set(source.lookups.flatMap((lookup) => lookup.masterFields.map((field) => field.code)));
    const fieldCodes = [
      ...source.fields.filter((field) => !lookupFieldCodes.has(field.code)).map((field) => field.code),
      ...source.lookups.map((lookup) => lookup.sourceField)
    ];
    const records = await getAllRecords(source.appId, query, fieldCodes, (count) => {
      setStatus(`取得中: ${sourceLabel} ${count}件`);
    });
    const rows = records.map((record) => {
      const row = Object.fromEntries(source.fields.map((field) => [field.code, fieldDisplayValue(record, field)]));
      source.lookups.forEach((lookup) => {
        if (!(lookup.sourceField in row)) {
          row[lookup.sourceField] = fieldDisplayValue(record, {
            code: lookup.sourceField,
            label: lookup.sourceField
          });
        }
      });
      return row;
    });
    await applySourceLookups(source, rows, setStatus);

    results.push({
      source,
      rows,
      periodStart: range.start,
      periodEnd: range.end,
      query
    });
  }

  return results;
}

async function validateSourceConfigs(config: PluginConfig, sources: SourceAppConfig[]): Promise<void> {
  const errors: string[] = [];
  const outputAppId = config.outputAppId || (config.mode === 'output' ? String(kintone.app.getId?.() || '') : '');
  const outputFieldFilters = sources.flatMap((source) =>
    source.filters
      .filter((filter) => filter.valueFrom === 'outputField')
      .map((filter) => ({ source, filter }))
  );

  let outputFieldCodes: Set<string> | null = null;
  if (outputFieldFilters.length) {
    if (!outputAppId) {
      errors.push('出力アプリの任意フィールドを抽出条件に使う場合は、紐づく帳票出力アプリIDを設定してください。');
    } else {
      try {
        outputFieldCodes = new Set((await getSourceAppFields(outputAppId)).fields.map((field) => field.code));
      } catch {
        errors.push(`帳票出力アプリ（アプリID: ${outputAppId}）が存在しない、または閲覧権限がありません。`);
      }
    }
  }

  for (const source of sources) {
    const sourceLabel = source.label || source.key || source.sheetName;
    if (!source.appId) {
      errors.push(`取得元アプリ「${sourceLabel}」のアプリIDが未設定です。`);
      continue;
    }

    let sourceFields: SourceFieldImportResult['fields'];
    try {
      sourceFields = (await getSourceAppFields(source.appId)).fields;
    } catch {
      errors.push(`取得元アプリ「${sourceLabel}」（アプリID: ${source.appId}）が存在しない、または閲覧権限がありません。`);
      continue;
    }

    const sourceFieldCodes = new Set(sourceFields.map((field) => field.code));
    if (outputFieldCodes) {
      source.filters
        .filter((filter) => filter.valueFrom === 'outputField')
        .forEach((filter) => {
          if (!filter.outputField) {
            errors.push(`取得元アプリ「${sourceLabel}」の抽出条件に出力アプリのフィールドコードが指定されていません。`);
          } else if (!outputFieldCodes?.has(filter.outputField)) {
            errors.push(`帳票出力アプリにフィールド「${filter.outputField}」が存在しません。`);
          }
        });
    }

    const lookupOutputFields = new Set(source.lookups.flatMap((lookup) => lookup.masterFields.map((field) => field.code)));
    const requiredSourceFields = [
      ...source.fields.filter((field) => !lookupOutputFields.has(field.code)).map((field) => field.code),
      ...source.filters.map((filter) => filter.field),
      ...source.sorts.map((sort) => sort.field),
      ...source.lookups.map((lookup) => lookup.sourceField)
    ];
    Array.from(new Set(requiredSourceFields.filter(Boolean))).forEach((fieldCode) => {
      if (!sourceFieldCodes.has(fieldCode)) {
        errors.push(`取得元アプリ「${sourceLabel}」にフィールド「${fieldCode}」が存在しません。`);
      }
    });

    for (const lookup of source.lookups) {
      const lookupLabel = `${sourceLabel} / マスタ参照 ${lookup.masterAppId}`;
      let masterFields: SourceFieldImportResult['fields'];
      try {
        masterFields = (await getSourceAppFields(lookup.masterAppId)).fields;
      } catch {
        errors.push(`マスタアプリ「${lookupLabel}」が存在しない、または閲覧権限がありません。`);
        continue;
      }

      const masterFieldCodes = new Set(masterFields.map((field) => field.code));
      [lookup.masterKeyField, ...lookup.masterFields.map((field) => field.code)].filter(Boolean).forEach((fieldCode) => {
        if (!masterFieldCodes.has(fieldCode)) {
          errors.push(`マスタアプリ「${lookupLabel}」にフィールド「${fieldCode}」が存在しません。`);
        }
      });
    }
  }

  if (errors.length) {
    throw new Error(['取得元アプリ設定の検証でエラーが見つかりました。', ...errors].join('\n'));
  }
}

async function applySourceLookups(
  source: SourceAppConfig,
  rows: Record<string, unknown>[],
  setStatus: (message: string) => void = () => undefined
): Promise<void> {
  for (const lookup of source.lookups) {
    const keys = Array.from(new Set(rows.map((row) => String(row[lookup.sourceField] ?? '')).filter(Boolean)));
    if (!keys.length || !lookup.masterFields.length) {
      continue;
    }

    setStatus(`マスタ参照中: ${source.label || source.key}`);
    const masterRecords = await fetchLookupMasterRecords(
      lookup.masterAppId,
      lookup.masterKeyField,
      lookup.masterFields,
      keys,
      (count) => setStatus(`マスタ参照中: ${source.label || source.key} ${count}件`)
    );
    const masterByKey = new Map(
      masterRecords.map((record) => [
        String(fieldDisplayValue(record, { code: lookup.masterKeyField, label: lookup.masterKeyField }) || ''),
        record
      ])
    );

    rows.forEach((row) => {
      const master = masterByKey.get(String(row[lookup.sourceField] ?? ''));
      if (!master) {
        return;
      }

      lookup.masterFields.forEach((field) => {
        row[field.code] = fieldDisplayValue(master, field);
      });
    });
  }
}

async function fetchLookupMasterRecords(
  appId: string,
  keyField: string,
  fields: SourceFieldConfig[],
  keys: string[],
  onProgress?: (count: number) => void
): Promise<KintoneRecord[]> {
  const records: KintoneRecord[] = [];
  const chunkSize = 100;
  for (let index = 0; index < keys.length; index += chunkSize) {
    const chunk = keys.slice(index, index + chunkSize);
    const query = buildInQuery(keyField, chunk);
    if (!query) {
      continue;
    }
    records.push(...(await getAllRecords(appId, query, [keyField, ...fields.map((field) => field.code)])));
    onProgress?.(records.length);
  }
  return records;
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

function applyOutputPeriodDefaults(config: PluginConfig, record: Record<string, any>, force = false): void {
  const baseDate = String(recordValue(record as KintoneRecord, config.outputBaseDateField) || '');
  if (!baseDate || (!config.outputPeriodStartField && !config.outputPeriodEndField)) {
    return;
  }

  const range = calculateDateRange('monthStartToBaseDate', baseDate);
  setRecordFieldValue(record, config.outputPeriodStartField, range.start, force);
  setRecordFieldValue(record, config.outputPeriodEndField, range.end, force);
}

function setRecordFieldValue(record: Record<string, any>, fieldCode: string, value: string, force: boolean): void {
  if (!fieldCode || !record[fieldCode]) {
    return;
  }
  if (force || !record[fieldCode].value) {
    record[fieldCode].value = value;
  }
}

async function saveOutputHistory(config: PluginConfig, context: ReportContext, fileName: string): Promise<void> {
  const recordId = kintone.app.record.getId?.();
  const appId = kintone.app.getId?.();
  const record: Record<string, { value: string }> = {};

  addHistoryField(record, config.outputBaseDateField, context.baseDate);
  addHistoryField(record, config.outputPeriodStartField, context.periodStart);
  addHistoryField(record, config.outputPeriodEndField, context.periodEnd);
  addHistoryField(record, config.outputExportedAtField, context.exportedAt);
  addHistoryField(record, config.outputExporterField, context.exporter);
  addHistoryField(record, config.outputFileNameField, fileName);
  addHistoryField(record, config.outputStatusField, '出力済み');
  addHistoryField(record, config.outputMemoField, buildOutputMemo(context, fileName));

  if (!recordId || !appId || !Object.keys(record).length) {
    return;
  }

  await kintone.api(kintone.api.url('/k/v1/record.json', true), 'PUT', {
    app: appId,
    id: recordId,
    record
  });
}

function buildOutputMemo(context: ReportContext, fileName: string): string {
  return [
    `帳票ID: ${context.reportId}`,
    `帳票名: ${context.reportName}`,
    `対象店舗: ${context.store}`,
    `基準日: ${context.baseDate}`,
    `対象期間: ${context.periodStart} - ${context.periodEnd}`,
    `出力日時: ${context.exportedAt}`,
    `出力者: ${context.exporter}`,
    `出力ファイル名: ${fileName}`
  ].join('\n');
}

function addHistoryField(record: Record<string, { value: string }>, fieldCode: string, value: string): void {
  if (fieldCode) {
    record[fieldCode] = { value };
  }
}

function formatDateTime(value: Date): string {
  const date = formatDate(value);
  const hours = String(value.getHours()).padStart(2, '0');
  const minutes = String(value.getMinutes()).padStart(2, '0');
  const seconds = String(value.getSeconds()).padStart(2, '0');
  return `${date}T${hours}:${minutes}:${seconds}`;
}

async function runWithStatus(
  button: HTMLButtonElement,
  status: HTMLElement,
  action: (setStatus: (message: string) => void) => Promise<void>
): Promise<void> {
  button.disabled = true;
  const setStatus = (message: string) => {
    status.textContent = message;
  };
  setStatus('処理中...');

  try {
    await action(setStatus);
    setStatus('完了しました。');
  } catch (error) {
    setStatus('エラー');
    showErrorDialog(errorMessage(error, '処理に失敗しました。'));
  } finally {
    button.disabled = false;
  }
}

function showErrorDialog(message: string): void {
  const dialog = document.createElement('dialog');
  if (typeof dialog.showModal !== 'function') {
    window.alert(message);
    return;
  }

  dialog.className = 'krp-error-dialog';
  dialog.innerHTML = `
    <form method="dialog" class="krp-error-dialog__body">
      <h2>処理に失敗しました</h2>
      <pre></pre>
      <button type="submit">閉じる</button>
    </form>
  `;
  const pre = dialog.querySelector('pre');
  if (pre) {
    pre.textContent = message;
  }
  dialog.addEventListener('close', () => dialog.remove());
  document.body.appendChild(dialog);
  dialog.showModal();
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
