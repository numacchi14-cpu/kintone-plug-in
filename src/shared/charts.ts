import JSZip from 'jszip';

// 完成版テンプレートに置いたグラフを出力ファイルへ残すための処理（v0.17.0）。
//
// ExcelJSはグラフを扱えず、読み込んだブックを書き出すとグラフ（xl/charts）と、グラフだけを置いた
// 描画（xl/drawings）が消える。openpyxlで作ったグラフは、読み込みの時点で例外になる。そこで、
// 1. ExcelJSへ渡す前に、グラフを含む描画をシートから外し（extractTemplateCharts）、
// 2. ExcelJSが書き出した出力ファイルへ、テンプレートの描画・グラフの部品をそのまま戻す
//    （restoreTemplateCharts）。
// グラフのないテンプレート（PL Managementのテンプレートすべて）では、どちらも何もしない。

const relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships';
const drawingRelationshipType = `${relationshipNamespace}/drawing`;
const officeDocumentRelationshipType = `${relationshipNamespace}/officeDocument`;
const chartGraphicDataPattern =
  /uri="http:\/\/schemas\.(?:openxmlformats\.org\/drawingml\/2006\/chart|microsoft\.com\/office\/drawing\/2014\/chartex)"/;
const pictureElementPattern = /<(?:[\w.-]+:)?pic[\s>/]/;
// ワークシートの子要素のうち、<drawing>より後ろに置く要素（OOXMLスキーマ CT_Worksheet の順序）。
const elementsAfterDrawing = new Set([
  'legacyDrawing',
  'legacyDrawingHF',
  'drawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst'
]);

export interface TemplateCharts {
  source: JSZip | null;
  drawings: Array<{ sheetName: string; drawingPath: string }>;
}

interface Relationship {
  element: string;
  id: string;
  type: string;
  target: string;
  external: boolean;
}

// テンプレートからグラフを含む描画をシートごと外したブックを返す。グラフがなければ元のバッファを
// そのまま返す（ExcelJSへ渡す内容を変えない）。
export async function extractTemplateCharts(templateBuffer: ArrayBuffer): Promise<{ buffer: ArrayBuffer; charts: TemplateCharts }> {
  const none = { buffer: templateBuffer, charts: { source: null, drawings: [] } };
  let source: JSZip;
  try {
    source = await JSZip.loadAsync(templateBuffer);
  } catch {
    // xlsxでないファイルは、従来どおりExcelJSの読み込みエラーとして扱う。
    return none;
  }

  const drawings: TemplateCharts['drawings'] = [];
  const removals: Array<{ sheetPath: string; relationship: Relationship }> = [];
  for (const [sheetName, sheetPath] of await worksheetPathsByName(source)) {
    for (const relationship of await readRelationships(source, sheetPath)) {
      if (relationship.type !== drawingRelationshipType || relationship.external) continue;
      const drawingPath = resolveTarget(sheetPath, relationship.target);
      const drawingXml = await source.file(drawingPath)?.async('string');
      if (!drawingXml || !chartGraphicDataPattern.test(drawingXml)) continue;
      if (pictureElementPattern.test(drawingXml)) {
        throw new Error(
          `テンプレートの「${sheetName}」シートに、グラフと画像が同じシート上に置かれています。グラフと画像を同じシートに置くことには対応していません。画像を別のシートへ移すか、削除してください。`
        );
      }
      drawings.push({ sheetName, drawingPath });
      removals.push({ sheetPath, relationship });
    }
  }
  if (!drawings.length) {
    return none;
  }

  const stripped = await JSZip.loadAsync(templateBuffer);
  for (const { sheetPath, relationship } of removals) {
    const sheetXml = await stripped.file(sheetPath)!.async('string');
    stripped.file(
      sheetPath,
      sheetXml.replace(/<(?:[\w.-]+:)?drawing\b(?:[^>"']|"[^"]*"|'[^']*')*\/>/g, (element) =>
        attributeValue(element, /(?:^|:)id$/) === relationship.id ? '' : element
      )
    );
    const relsPath = relationshipsPath(sheetPath);
    const relsXml = await stripped.file(relsPath)!.async('string');
    stripped.file(relsPath, relsXml.replace(relationship.element, ''));
  }
  // ExcelJSは描画の部品をシートと無関係に読み込むため、外した描画の部品も渡さない。
  for (const { drawingPath } of drawings) {
    stripped.remove(drawingPath);
    stripped.remove(relationshipsPath(drawingPath));
  }

  return { buffer: await stripped.generateAsync({ type: 'arraybuffer' }), charts: { source, drawings } };
}

// ExcelJSが書き出した出力ファイルへ、テンプレートの描画とグラフを戻す。シートの対応はシート名で取る
// （ExcelJSはsheetN.xmlの番号を振り直すことがある）。部品名は出力ファイルの既存の部品と重ならないように付け直す。
export async function restoreTemplateCharts(outputBuffer: ArrayBuffer, charts: TemplateCharts): Promise<ArrayBuffer> {
  const source = charts.source;
  if (!source || !charts.drawings.length) {
    return outputBuffer;
  }

  const output = await JSZip.loadAsync(outputBuffer);
  const outputSheets = await worksheetPathsByName(output);
  const templateTypes = parseContentTypes(await source.file('[Content_Types].xml')?.async('string') ?? '');
  const usedPartNames = new Set(Object.keys(output.files).map((name) => name.toLowerCase()));
  const copiedParts = new Map<string, string>();
  const overrides: string[] = [];

  const copyPart = async (sourcePath: string): Promise<string> => {
    const copied = copiedParts.get(sourcePath);
    if (copied) return copied;
    const file = source.file(sourcePath);
    if (!file) {
      throw new Error(`テンプレートのグラフが参照する部品「${sourcePath}」が見つかりません。テンプレートをExcelで開いて保存し直してください。`);
    }
    const targetPath = allocatePartName(sourcePath, usedPartNames);
    copiedParts.set(sourcePath, targetPath);
    output.file(targetPath, await file.async('uint8array'));
    const contentType = templateTypes.overrides.get(sourcePath.toLowerCase()) ?? templateTypes.defaults.get(extensionOf(sourcePath));
    if (contentType) {
      overrides.push(`<Override PartName="/${escapeXml(targetPath)}" ContentType="${escapeXml(contentType)}"/>`);
    }

    // グラフのスタイル・色（chartStyle・colors）など、この部品がさらに参照する部品もコピーし、参照先を付け替える。
    const relsXml = await source.file(relationshipsPath(sourcePath))?.async('string');
    if (relsXml) {
      let rewritten = relsXml;
      for (const relationship of parseRelationships(relsXml)) {
        if (relationship.external) continue;
        const childPath = await copyPart(resolveTarget(sourcePath, relationship.target));
        rewritten = rewritten.replace(
          relationship.element,
          setAttribute(relationship.element, 'Target', relativeTarget(targetPath, childPath))
        );
      }
      output.file(relationshipsPath(targetPath), rewritten);
    }
    return targetPath;
  };

  for (const { sheetName, drawingPath } of charts.drawings) {
    const sheetPath = outputSheets.get(sheetName);
    if (!sheetPath) {
      throw new Error(`出力ファイルに「${sheetName}」シートが見つからないため、テンプレートのグラフを出力できません。`);
    }
    const sheetXml = await output.file(sheetPath)!.async('string');
    const children = topLevelChildren(sheetXml);
    if (children.elements.some((child) => child.name === 'drawing')) {
      throw new Error(
        `出力ファイルの「${sheetName}」シートに画像などの描画があるため、テンプレートのグラフを出力できません。グラフと画像を同じシートに置くことには対応していません。`
      );
    }

    const newDrawingPath = await copyPart(drawingPath);
    const relsPath = relationshipsPath(sheetPath);
    const relsXml =
      (await output.file(relsPath)?.async('string')) ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${packageRelationshipNamespace}"></Relationships>`;
    const relationshipId = uniqueRelationshipId(parseRelationships(relsXml));
    const relationshipElement = `<Relationship Id="${relationshipId}" Type="${drawingRelationshipType}" Target="${escapeXml(relativeTarget(sheetPath, newDrawingPath))}"/>`;
    output.file(
      relsPath,
      /<Relationships\b[^>]*\/>/.test(relsXml)
        ? relsXml.replace(/<Relationships\b([^>]*)\/>/, `<Relationships$1>${relationshipElement}</Relationships>`)
        : relsXml.replace('</Relationships>', `${relationshipElement}</Relationships>`)
    );
    output.file(sheetPath, insertDrawingElement(sheetXml, children, relationshipId));
  }

  const contentTypesXml = await output.file('[Content_Types].xml')!.async('string');
  output.file('[Content_Types].xml', contentTypesXml.replace('</Types>', `${overrides.join('')}</Types>`));
  return output.generateAsync({ type: 'arraybuffer' });
}

// ワークシートの<drawing r:id>を、スキーマ上の正しい位置（pageSetup・headerFooter・rowBreaksなどの後、
// legacyDrawing・tableParts・extLstなどの前）へ入れる。
function insertDrawingElement(sheetXml: string, children: ReturnType<typeof topLevelChildren>, relationshipId: string): string {
  let rootTag = children.rootTag;
  const declarations = [...rootTag.matchAll(/\sxmlns:([\w.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)];
  let prefix = declarations.find((match) => (match[2] ?? match[3]) === relationshipNamespace)?.[1];
  if (!prefix) {
    const declared = new Set(declarations.map((match) => match[1]));
    prefix = declared.has('r') ? 'krpr' : 'r';
    rootTag = rootTag.replace(/(\/?>)$/, ` xmlns:${prefix}="${relationshipNamespace}"$1`);
  }
  const elementPrefix = children.rootName.includes(':') ? `${children.rootName.split(':')[0]}:` : '';
  const drawingElement = `<${elementPrefix}drawing ${prefix}:id="${relationshipId}"/>`;
  const insertAt = children.elements.find((child) => elementsAfterDrawing.has(child.name))?.index ?? children.closeIndex;
  const withDrawing = `${sheetXml.slice(0, insertAt)}${drawingElement}${sheetXml.slice(insertAt)}`;
  return withDrawing.slice(0, children.rootIndex) + rootTag + withDrawing.slice(children.rootIndex + children.rootTag.length);
}

// ルート要素の開始タグと、直下の子要素の位置・名前（接頭辞なし）を返す。
function topLevelChildren(xml: string) {
  const tagPattern = /<(\/?)([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  const elements: Array<{ name: string; index: number }> = [];
  let depth = 0;
  let rootTag = '';
  let rootName = '';
  let rootIndex = -1;
  let closeIndex = -1;
  for (const match of xml.matchAll(tagPattern)) {
    const [tag, closing, name, , selfClosing] = match;
    const index = match.index ?? 0;
    if (closing) {
      depth -= 1;
      if (depth === 0) closeIndex = index;
      continue;
    }
    if (depth === 0) {
      rootTag = tag;
      rootName = name;
      rootIndex = index;
    } else if (depth === 1) {
      elements.push({ name: name.replace(/^[\w.-]+:/, ''), index });
    }
    if (!selfClosing) depth += 1;
  }
  if (rootIndex < 0 || closeIndex < 0) {
    throw new Error('出力ファイルのワークシートを解析できないため、テンプレートのグラフを出力できません。');
  }
  return { rootTag, rootName, rootIndex, closeIndex, elements };
}

// ブック内のワークシートの部品名を、シート名から引けるようにする。
async function worksheetPathsByName(zip: JSZip): Promise<Map<string, string>> {
  const rootRelationships = await readRelationships(zip, '');
  const workbookTarget = rootRelationships.find((relationship) => relationship.type === officeDocumentRelationshipType)?.target;
  const workbookPath = workbookTarget ? resolveTarget('', workbookTarget) : 'xl/workbook.xml';
  const workbookXml = (await zip.file(workbookPath)?.async('string')) ?? '';
  const relationships = new Map((await readRelationships(zip, workbookPath)).map((relationship) => [relationship.id, relationship]));

  const sheets = new Map<string, string>();
  for (const [element] of workbookXml.matchAll(/<(?:[\w.-]+:)?sheet\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/g)) {
    const name = attributeValue(element, /^name$/);
    const relationship = relationships.get(attributeValue(element, /:id$/) ?? '');
    if (name == null || !relationship) continue;
    sheets.set(decodeXml(name), resolveTarget(workbookPath, relationship.target));
  }
  return sheets;
}

async function readRelationships(zip: JSZip, partPath: string): Promise<Relationship[]> {
  const xml = await zip.file(relationshipsPath(partPath))?.async('string');
  return xml ? parseRelationships(xml) : [];
}

function parseRelationships(xml: string): Relationship[] {
  return [...xml.matchAll(/<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*?\/>/g)].map(([element]) => ({
    element,
    id: attributeValue(element, /^Id$/) ?? '',
    type: attributeValue(element, /^Type$/) ?? '',
    target: decodeXml(attributeValue(element, /^Target$/) ?? ''),
    external: attributeValue(element, /^TargetMode$/) === 'External'
  }));
}

function parseContentTypes(xml: string) {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const [element] of xml.matchAll(/<(Override|Default)\b(?:[^>"']|"[^"]*"|'[^']*')*\/>/g)) {
    const contentType = decodeXml(attributeValue(element, /^ContentType$/) ?? '');
    const partName = attributeValue(element, /^PartName$/);
    const extension = attributeValue(element, /^Extension$/);
    if (partName != null) overrides.set(decodeXml(partName).replace(/^\//, '').toLowerCase(), contentType);
    if (extension != null) defaults.set(decodeXml(extension).toLowerCase(), contentType);
  }
  return { overrides, defaults };
}

function uniqueRelationshipId(relationships: Relationship[]): string {
  const ids = new Set(relationships.map((relationship) => relationship.id));
  let number = relationships.length + 1;
  while (ids.has(`rId${number}`)) number += 1;
  return `rId${number}`;
}

// xl/charts/chart1.xml → 出力ファイルで未使用の xl/charts/chartN.xml。
function allocatePartName(sourcePath: string, usedPartNames: Set<string>): string {
  const match = sourcePath.match(/^(.*\/)?([^/]*?)(\d*)(\.[^./]*)?$/);
  const directory = match?.[1] ?? '';
  const stem = match?.[2] ?? 'part';
  const extension = match?.[4] ?? '';
  let number = 1;
  while (usedPartNames.has(`${directory}${stem}${number}${extension}`.toLowerCase())) number += 1;
  const partName = `${directory}${stem}${number}${extension}`;
  usedPartNames.add(partName.toLowerCase());
  return partName;
}

function relationshipsPath(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

// 関係の参照先を部品名へ直す。openpyxlは絶対パス（/xl/...）、Excelは相対パス（../...）で書く。
function resolveTarget(fromPart: string, target: string): string {
  const base = target.startsWith('/') ? [] : fromPart.split('/').slice(0, -1);
  for (const segment of target.replace(/^\//, '').split('/')) {
    if (segment === '..') base.pop();
    else if (segment && segment !== '.') base.push(segment);
  }
  return base.join('/');
}

function relativeTarget(fromPart: string, toPart: string): string {
  const from = fromPart.split('/').slice(0, -1);
  const to = toPart.split('/');
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common += 1;
  return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/');
}

function extensionOf(partPath: string): string {
  const match = partPath.match(/\.([^./]+)$/);
  return match ? match[1].toLowerCase() : '';
}

function attributeValue(element: string, namePattern: RegExp): string | undefined {
  for (const match of element.matchAll(/\s([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    if (namePattern.test(match[1])) return match[2] ?? match[3];
  }
  return undefined;
}

function setAttribute(element: string, name: string, value: string): string {
  return element.replace(new RegExp(`(\\s${name}\\s*=\\s*)(?:"[^"]*"|'[^']*')`), `$1"${escapeXml(value)}"`);
}

function decodeXml(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_, entity: string) => {
    if (entity[0] === '#') {
      return String.fromCodePoint(entity[1] === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1)));
    }
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[entity] ?? '';
  });
}

function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
