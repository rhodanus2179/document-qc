import { RULES } from './rules.js';

const ERROR_VALUES = ['#REF!', '#VALUE!', '#NAME?', '#N/A', '#DIV/0!', '#NUM!', '#NULL!'];

export function runRules(model, enabledIds = new Set(RULES.map(r => r.id)), onProgress = () => {}) {
  const findings = [];
  const applicable = RULES.filter(r => enabledIds.has(r.id) && r.appliesTo.includes(model.kind));
  applicable.forEach((rule, index) => {
    const push = item => findings.push({
      ruleId: rule.id,
      category: rule.category,
      severity: rule.severity,
      title: rule.title,
      matched: '—', suggestion: '—', context: '',
      ...item
    });
    runOne(rule, model, push);
    onProgress(60 + Math.round(((index + 1) / Math.max(1, applicable.length)) * 35), `ルール ${index + 1}/${applicable.length}: ${rule.id}`);
  });
  return findings;
}

function runOne(rule, model, push) {
  switch (rule.engine) {
    case 'nbsp': return forParagraphMatch(model, /\u00a0/g, push, '不可視のNBSP（U+00A0）が含まれています。', '通常の空白または不要なら削除');
    case 'repeated-space': return forParagraphMatch(model, / {2,}/g, push, '半角空白が連続しています。', '必要なければ1文字に統一');
    case 'repeated-punctuation': return forParagraphMatch(model, /(、、+|。。+|，，+|．．+)/g, push, '句読点が連続しています。', '重複を確認して修正');
    case 'punctuation-mix': return punctuationMix(model, push);
    case 'register-mix': return registerMix(model, push);
    case 'empty-table': return model.tables.filter(t => !t.text).forEach(t => push({ location: `表 ${t.index}`, message: `${t.rows}行×${t.cols}列の空の表があります。`, suggestion: '不要であれば表を削除' }));
    case 'red-text': return redText(model, push);
    case 'comments': if (model.hasComments) push({ location: '文書全体', message: 'コメントが残っています。', suggestion: '必要性を確認し、成果品では削除を検討' }); return;
    case 'tracked-changes': if (model.hasTrackedChanges) push({ location: '文書全体', message: '変更履歴が残っています。', suggestion: '変更履歴を承諾または削除して確認' }); return;
    case 'broken-reference-text': return forParagraphMatch(model, /(エラー!\s*参照元が見つかりません。?|Error!\s*Reference source not found\.?)/gi, push, '相互参照のエラー文字列があります。', '参照先フィールドを確認');
    case 'missing-caption': return missingCaptions(model, push);
    case 'caption-gap': return captionGaps(model, push);
    case 'error-cell': return excelErrorCells(model, push);
    case 'broken-formula-ref': return eachCell(model, ({ sheet, cell }) => { if (cell.formula?.includes('#REF!')) push(cellFinding(sheet, cell, '数式に #REF! が含まれています。', '参照先を修正')); });
    case 'hidden-sheet': return model.sheets.filter(s => s.state !== 'visible').forEach(s => push({ location: `シート「${s.name}」`, matched: s.state, message: `シートが ${s.state} 状態です。`, suggestion: '意図した非表示か確認' }));
    case 'hidden-row': return model.sheets.forEach(s => { if (s.hiddenRows.length) push({ location: `シート「${s.name}」`, matched: summarizeNumbers(s.hiddenRows), message: `非表示行が ${s.hiddenRows.length} 行あります。`, suggestion: '意図した非表示か確認' }); });
    case 'hidden-col': return model.sheets.forEach(s => { if (s.hiddenCols.length) push({ location: `シート「${s.name}」`, matched: s.hiddenCols.map(columnName).slice(0, 12).join(', '), message: `非表示列が ${s.hiddenCols.length} 列あります。`, suggestion: '意図した非表示か確認' }); });
    case 'external-link': if (model.hasExternalLinks) push({ location: 'ブック全体', message: '外部リンクのOOXMLパーツが含まれています。', suggestion: 'リンク先と必要性を確認' }); return;
    case 'formula-value-outlier': return formulaValueOutliers(model, push);
    case 'formula-pattern-outlier': return formulaPatternOutliers(model, push);
    case 'self-reference': return selfReferences(model, push);
  }
}

function forParagraphMatch(model, regex, push, message, suggestion) {
  for (const p of model.paragraphs) {
    regex.lastIndex = 0;
    let m;
    while ((m = regex.exec(p.text))) {
      push({ location: paragraphLocation(p), matched: visibleMatch(m[0]), message, suggestion, context: excerpt(p.text, m.index, m[0].length) });
      if (!regex.global) break;
    }
  }
}

function punctuationMix(model, push) {
  const counts = { '、': 0, '，': 0, '。': 0, '．': 0 };
  for (const ch of model.text) if (ch in counts) counts[ch]++;
  for (const [a, b] of [['、', '，'], ['。', '．']]) {
    if (counts[a] && counts[b]) {
      const preferred = counts[a] >= counts[b] ? a : b;
      push({ location: '文書全体', matched: `${a}:${counts[a]} / ${b}:${counts[b]}`, message: `「${a}」と「${b}」が混在しています。文書内の多数派は「${preferred}」です。`, suggestion: `意図に応じて「${preferred}」へ統一` });
    }
  }
}

function registerMix(model, push) {
  let polite = 0, plain = 0;
  for (const p of model.paragraphs) {
    if (/(です|ます|でした|ません)[。．！？!?]?\s*$/.test(p.text)) polite++;
    if (/(である|だった|となる|とする)[。．！？!?]?\s*$/.test(p.text)) plain++;
  }
  if (polite && plain) {
    const preferred = polite >= plain ? '敬体（です・ます調）' : '常体（である調）';
    push({ location: '文書全体', matched: `敬体 ${polite} / 常体 ${plain}`, message: `文体が混在しています。簡易判定では ${preferred} が多数です。`, suggestion: '引用等を除き、文書方針に合わせて確認' });
  }
}

function redText(model, push) {
  for (const p of model.paragraphs) for (const run of p.runs) {
    if (run.color === 'FF0000' || run.color === 'F00') push({ location: paragraphLocation(p), matched: run.text || '(空文字)', message: '直接指定の赤色文字が残っています。', suggestion: '校正跡か意図した強調か確認', context: excerpt(p.text, Math.max(0, p.text.indexOf(run.text)), run.text.length) });
  }
}

function missingCaptions(model, push) {
  const captionKeys = new Set();
  const references = [];
  for (const p of model.paragraphs) {
    const caption = p.text.match(/^\s*(図|表)\s*([0-9０-９]+(?:\s*[-－―ー]\s*[0-9０-９]+)+)/);
    if (caption) captionKeys.add(normalizeFigureKey(caption[1], caption[2]));
    const re = /(図|表)\s*([0-9０-９]+(?:\s*[-－―ー]\s*[0-9０-９]+)+)/g;
    let m;
    while ((m = re.exec(p.text))) references.push({ key: normalizeFigureKey(m[1], m[2]), raw: m[0], p });
  }
  for (const ref of references) {
    if (!captionKeys.has(ref.key)) push({ location: paragraphLocation(ref.p), matched: ref.raw, message: `本文参照「${ref.raw}」に対応するキャプションが見つかりません。`, suggestion: '図表番号と参照先を確認', context: excerpt(ref.p.text, ref.p.text.indexOf(ref.raw), ref.raw.length) });
  }
}

function captionGaps(model, push) {
  const groups = new Map();
  for (const p of model.paragraphs) {
    const m = p.text.match(/^\s*(図|表)\s*([0-9０-９]+)\s*[-－―ー]\s*([0-9０-９]+)/);
    if (!m) continue;
    const kind = m[1], major = toAsciiDigits(m[2]), minor = Number(toAsciiDigits(m[3]));
    const key = `${kind}:${major}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ minor, p, raw: m[0] });
  }
  for (const [key, items] of groups) {
    items.sort((a, b) => a.minor - b.minor);
    for (let i = 1; i < items.length; i++) {
      if (items[i].minor > items[i - 1].minor + 1) push({ location: paragraphLocation(items[i].p), matched: items[i].raw.trim(), message: `${key.replace(':', '')} 系列の番号が ${items[i - 1].minor} から ${items[i].minor} に飛んでいます。`, suggestion: '欠番が意図したものか確認' });
    }
  }
}

function excelErrorCells(model, push) {
  eachCell(model, ({ sheet, cell }) => {
    const combined = `${cell.value ?? ''} ${cell.formula ?? ''}`;
    const hit = ERROR_VALUES.find(e => combined.includes(e));
    if (hit) push(cellFinding(sheet, cell, `${hit} を検出しました。`, '原因となる数式・参照を確認', hit));
  });
}

function formulaValueOutliers(model, push) {
  for (const sheet of model.sheets) {
    const map = cellMap(sheet.cells);
    for (const cell of sheet.cells) {
      if (cell.formula || cell.value === '') continue;
      const { row, col } = parseCellRef(cell.ref);
      if (!row || !col) continue;
      const vertical = hasFormula(map, row - 1, col) && hasFormula(map, row + 1, col);
      const horizontal = hasFormula(map, row, col - 1) && hasFormula(map, row, col + 1);
      if (vertical || horizontal) push(cellFinding(sheet, cell, '周囲が数式セルですが、このセルだけ直値です。', '意図した上書きか確認', String(cell.value)));
    }
  }
}

function formulaPatternOutliers(model, push) {
  for (const sheet of model.sheets) {
    const map = cellMap(sheet.cells);
    for (const cell of sheet.cells) {
      if (!cell.formula) continue;
      const { row, col } = parseCellRef(cell.ref);
      if (!row || !col) continue;
      const above = map.get(key(row - 1, col)), below = map.get(key(row + 1, col));
      const left = map.get(key(row, col - 1)), right = map.get(key(row, col + 1));
      const currentNorm = normalizeFormula(cell.formula, row, col);
      let expected = null;
      if (above?.formula && below?.formula) {
        const a = normalizeFormula(above.formula, row - 1, col), b = normalizeFormula(below.formula, row + 1, col);
        if (a === b) expected = a;
      }
      if (!expected && left?.formula && right?.formula) {
        const a = normalizeFormula(left.formula, row, col - 1), b = normalizeFormula(right.formula, row, col + 1);
        if (a === b) expected = a;
      }
      if (expected && currentNorm !== expected) push(cellFinding(sheet, cell, '周囲と相対参照パターンが異なる数式です。', 'コピーずれ・例外式か確認', `=${cell.formula}`));
    }
  }
}

function selfReferences(model, push) {
  eachCell(model, ({ sheet, cell }) => {
    if (!cell.formula) return;
    const ref = parseCellRef(cell.ref);
    if (!ref.row || !ref.col) return;
    const re = new RegExp(`(^|[^A-Z0-9_])\\$?${columnName(ref.col)}\\$?${ref.row}(?![A-Z0-9_])`, 'i');
    if (re.test(cell.formula)) push(cellFinding(sheet, cell, '数式が自身のセルを直接参照している可能性があります。', '循環参照が意図したものか確認', `=${cell.formula}`));
  });
}

function eachCell(model, fn) { for (const sheet of model.sheets) for (const cell of sheet.cells) fn({ sheet, cell }); }
function cellFinding(sheet, cell, message, suggestion, matched = null) { return { location: `シート「${sheet.name}」 ${cell.ref}`, matched: matched ?? String(cell.value || cell.formula || '—'), message, suggestion, context: cell.formula ? `=${cell.formula}` : String(cell.value ?? '') }; }
function paragraphLocation(p) { return `${p.pageHint ? `保存時ページ目安 ${p.pageHint} / ` : ''}第${p.index}段落`; }
function visibleMatch(s) { return s === '\u00a0' ? 'U+00A0' : s.replace(/\u00a0/g, '⍽'); }
function excerpt(text, start, length) { const a = Math.max(0, start - 24), b = Math.min(text.length, start + length + 36); return `${a ? '…' : ''}${text.slice(a, b)}${b < text.length ? '…' : ''}`; }
function toAsciiDigits(s) { return s.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)); }
function normalizeFigureKey(kind, num) { return `${kind}:${toAsciiDigits(num).replace(/\s/g, '').replace(/[－―ー]/g, '-')}`; }
function summarizeNumbers(nums) { const shown = nums.slice(0, 15).join(', '); return nums.length > 15 ? `${shown} …` : shown; }
function columnName(num) { let s = ''; for (let n = num; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; }
function parseCellRef(ref) { const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref); if (!m) return { row: 0, col: 0 }; let col = 0; for (const ch of m[1].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64; return { row: Number(m[2]), col }; }
function key(row, col) { return `${row}:${col}`; }
function cellMap(cells) { const m = new Map(); for (const c of cells) { const p = parseCellRef(c.ref); if (p.row && p.col) m.set(key(p.row, p.col), c); } return m; }
function hasFormula(map, row, col) { return Boolean(row > 0 && col > 0 && map.get(key(row, col))?.formula); }

function normalizeFormula(formula, baseRow, baseCol) {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/gi, (full, absCol, letters, absRow, rowText) => {
    const refCol = parseCellRef(`${letters}${rowText}`).col;
    const refRow = Number(rowText);
    const c = absCol ? `C$${refCol}` : `C[${refCol - baseCol}]`;
    const r = absRow ? `R$${refRow}` : `R[${refRow - baseRow}]`;
    return `${r}${c}`;
  });
}
