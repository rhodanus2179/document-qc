import { readZip, decodeEntry, normalizePath } from './zip.js';

const NS = '*';

export async function parseOfficeFile(file, onProgress = () => {}) {
  const filename = file?.name || '';
  validateFilename(filename);
  onProgress(15, 'ファイルをメモリへ読み込んでいます');
  const buffer = await file.arrayBuffer();
  return parseOfficeBuffer(buffer, filename, onProgress);
}

export async function parseOfficeBuffer(buffer, filename, onProgress = () => {}) {
  validateFilename(filename);
  const lower = filename.toLowerCase();
  const arrayBuffer = normalizeArrayBuffer(buffer);
  onProgress(35, 'OOXMLパッケージを展開しています');
  const entries = await readZip(arrayBuffer);
  onProgress(55, '文書構造を解析しています');
  if (lower.endsWith('.docx')) return parseDocx(entries, filename);
  return parseXlsx(entries, filename);
}

function validateFilename(filename) {
  const lower = String(filename || '').toLowerCase();
  if (!lower.endsWith('.docx') && !lower.endsWith('.xlsx')) throw new Error('対応形式は .docx と .xlsx です。');
}

function normalizeArrayBuffer(value) {
  if (value instanceof ArrayBuffer) return value;
  if (ArrayBuffer.isView(value)) return value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength);
  throw new Error('Office文書データを ArrayBuffer として読み込めません。');
}

function parseXml(text, label) {
  if (!text) return null;
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error(`${label} のXMLを解析できません。`);
  return doc;
}

function all(root, localName) { return root ? Array.from(root.getElementsByTagNameNS(NS, localName)) : []; }
function children(root, localName) { return root ? Array.from(root.children).filter(el => el.localName === localName) : []; }
function first(root, localName) { return all(root, localName)[0] || null; }
function attr(node, localName) {
  if (!node) return null;
  for (const a of Array.from(node.attributes || [])) if (a.localName === localName) return a.value;
  return null;
}
function textOf(root) { return all(root, 't').map(n => n.textContent || '').join(''); }

function paragraphText(p) {
  let out = '';
  const walk = node => {
    for (const child of Array.from(node.childNodes || [])) {
      if (child.nodeType !== 1) continue;
      if (child.localName === 't' || child.localName === 'delText') out += child.textContent || '';
      else if (child.localName === 'tab') out += '\t';
      else if (child.localName === 'br') out += '\n';
      else walk(child);
    }
  };
  walk(p);
  return out;
}

function parseDocx(entries, filename) {
  const documentXml = parseXml(decodeEntry(entries, 'word/document.xml'), 'word/document.xml');
  if (!documentXml) throw new Error('Word本文が見つかりません。');
  const paragraphs = [];
  let page = 1;
  for (const [index, p] of all(documentXml, 'p').entries()) {
    const pPr = children(p, 'pPr')[0];
    const pStyle = pPr ? first(pPr, 'pStyle') : null;
    const styleId = attr(pStyle, 'val') || '';
    const runs = children(p, 'r').map(r => {
      const rPr = children(r, 'rPr')[0];
      const color = attr(rPr ? first(rPr, 'color') : null, 'val');
      const sizeHalfPoints = Number(attr(rPr ? first(rPr, 'sz') : null, 'val') || 0);
      return {
        text: paragraphText(r),
        color: color && color !== 'auto' ? color.toUpperCase() : null,
        sizePt: sizeHalfPoints ? sizeHalfPoints / 2 : null,
        bold: Boolean(rPr && first(rPr, 'b')),
        italic: Boolean(rPr && first(rPr, 'i'))
      };
    });
    const text = paragraphText(p);
    paragraphs.push({ index: index + 1, pageHint: page, styleId, text, runs });
    const explicitBreaks = all(p, 'br').filter(br => attr(br, 'type') === 'page').length;
    const renderedBreaks = all(p, 'lastRenderedPageBreak').length;
    page += explicitBreaks + renderedBreaks;
  }

  const tables = all(documentXml, 'tbl').map((tbl, index) => {
    const rows = children(tbl, 'tr');
    const matrix = rows.map(tr => children(tr, 'tc').map(tc => textOf(tc)));
    return { index: index + 1, rows: matrix.length, cols: Math.max(0, ...matrix.map(r => r.length)), cells: matrix, text: matrix.flat().join(' ').trim() };
  });

  const hasComments = entries.has('word/comments.xml') || all(documentXml, 'commentReference').length > 0;
  const hasTrackedChanges = all(documentXml, 'ins').length > 0 || all(documentXml, 'del').length > 0 || all(documentXml, 'moveFrom').length > 0 || all(documentXml, 'moveTo').length > 0;

  return {
    kind: 'word', filename,
    paragraphs, tables, hasComments, hasTrackedChanges,
    text: paragraphs.map(p => p.text).join('\n'),
    packageEntries: new Set(entries.keys())
  };
}

function parseXlsx(entries, filename) {
  const workbook = parseXml(decodeEntry(entries, 'xl/workbook.xml'), 'xl/workbook.xml');
  if (!workbook) throw new Error('Excelワークブックが見つかりません。');
  const rels = parseXml(decodeEntry(entries, 'xl/_rels/workbook.xml.rels'), 'workbook.xml.rels');
  const relMap = new Map();
  for (const rel of all(rels, 'Relationship')) relMap.set(attr(rel, 'Id'), attr(rel, 'Target'));

  const sharedStringsDoc = parseXml(decodeEntry(entries, 'xl/sharedStrings.xml'), 'xl/sharedStrings.xml');
  const sharedStrings = sharedStringsDoc ? all(sharedStringsDoc, 'si').map(si => textOf(si)) : [];

  const sheets = [];
  for (const sheetNode of all(workbook, 'sheet')) {
    const name = attr(sheetNode, 'name') || 'Sheet';
    const state = attr(sheetNode, 'state') || 'visible';
    const relId = attr(sheetNode, 'id');
    const target = relMap.get(relId);
    if (!target) continue;
    const sheetPath = normalizePath(target.startsWith('/') ? target.slice(1) : `xl/${target}`);
    const sheetDoc = parseXml(decodeEntry(entries, sheetPath), sheetPath);
    if (!sheetDoc) continue;
    const hiddenRows = all(sheetDoc, 'row').filter(row => attr(row, 'hidden') === '1').map(row => Number(attr(row, 'r'))).filter(Boolean);
    const hiddenCols = [];
    for (const col of all(sheetDoc, 'col')) {
      if (attr(col, 'hidden') !== '1') continue;
      const min = Number(attr(col, 'min')), max = Number(attr(col, 'max'));
      for (let i = min; i <= max && i <= min + 500; i++) hiddenCols.push(i);
    }
    const cells = [];
    for (const c of all(sheetDoc, 'c')) {
      const ref = attr(c, 'r') || '';
      const type = attr(c, 't') || '';
      const formula = children(c, 'f')[0]?.textContent || null;
      const v = children(c, 'v')[0]?.textContent ?? null;
      let value = v;
      if (type === 's' && v !== null) value = sharedStrings[Number(v)] ?? v;
      else if (type === 'inlineStr') value = textOf(children(c, 'is')[0]);
      else if (type === 'b') value = v === '1' ? 'TRUE' : 'FALSE';
      cells.push({ ref, type, formula, value: value ?? '' });
    }
    sheets.push({ name, state, path: sheetPath, cells, hiddenRows, hiddenCols });
  }

  return {
    kind: 'excel', filename, sheets,
    hasExternalLinks: Array.from(entries.keys()).some(k => k.startsWith('xl/externalLinks/')),
    packageEntries: new Set(entries.keys())
  };
}
