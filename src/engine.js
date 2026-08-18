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
    case 'literal-replacement': return literalReplacements(rule, model, push);
    case 'kw-case': return scanContentMatch(model, /(?<![A-Za-z])(?:KW|kw)(?![A-Za-z])/g, push, '単位 kW の大文字小文字を確認してください。', 'kW');
    case 'approx-redundancy': return approximateRedundancy(model, push);
    case 'each-redundancy': return eachRedundancy(model, push);
    case 'double-negative': return scanContentMatch(model, /(ないわけではない|なくはない|ないこともない)/g, push, '二重否定になっています。意図したニュアンスか確認してください。', '可能であれば肯定表現に言い換え');
    case 'start-time-yori': return startTimeYori(model, push);
    case 'percent-width-mix': return notationMix(model, push, '％', '%', 'パーセント記号');
    case 'co2-mix': return notationMix(model, push, 'CO2', 'CO₂', 'CO2表記');
    case 'slash-width-mix': return notationMix(model, push, '／', '/', 'スラッシュ');
    case 'year-format-mix': return yearFormatMix(model, push);
    case 'ton-unit-mix': return tonUnitMix(model, push);
    case 'tco2-mix': return notationMix(model, push, 't-CO2', 'tCO2', 'CO2質量単位');
    case 'nth-day': return nthDay(model, push);
    case 'fullwidth-unit': return fullwidthUnit(model, push);

    case 'empty-table': return model.tables.filter(t => !t.text).forEach(t => push({ location: `表 ${t.index}`, message: `${t.rows}行×${t.cols}列の空の表があります。`, suggestion: '不要であれば表を削除' }));
    case 'red-text': return redText(model, push);
    case 'comments': if (model.hasComments) push({ location: '文書全体', message: 'コメントが残っています。', suggestion: '必要性を確認し、成果品では削除を検討' }); return;
    case 'tracked-changes': if (model.hasTrackedChanges) push({ location: '文書全体', message: '変更履歴が残っています。', suggestion: '変更履歴を承諾または削除して確認' }); return;
    case 'broken-reference-text': return forParagraphMatch(model, /(エラー!\s*参照元が見つかりません。?|Error!\s*Reference source not found\.?)/gi, push, '相互参照のエラー文字列があります。', '参照先フィールドを確認');
    case 'missing-caption': return missingCaptions(model, push);
    case 'caption-gap': return captionGaps(model, push);
    case 'caption-duplicate': return captionDuplicates(model, push);
    case 'heading-edge-space': return headingEdgeSpaces(model, push);
    case 'paragraph-trailing-space': return paragraphTrailingSpaces(model, push);

    case 'error-cell': return excelErrorCells(model, push);
    case 'broken-formula-ref': return eachCell(model, ({ sheet, cell }) => { if (cell.formula?.includes('#REF!')) push(cellFinding(sheet, cell, '数式に #REF! が含まれています。', '参照先を修正')); });
    case 'hidden-sheet': return model.sheets.filter(s => s.state !== 'visible').forEach(s => push({ location: `シート「${s.name}」`, matched: s.state, message: `シートが ${s.state} 状態です。`, suggestion: '意図した非表示か確認' }));
    case 'hidden-row': return model.sheets.forEach(s => { if (s.hiddenRows.length) push({ location: `シート「${s.name}」`, matched: summarizeNumbers(s.hiddenRows), message: `非表示行が ${s.hiddenRows.length} 行あります。`, suggestion: '意図した非表示か確認' }); });
    case 'hidden-col': return model.sheets.forEach(s => { if (s.hiddenCols.length) push({ location: `シート「${s.name}」`, matched: s.hiddenCols.map(columnName).slice(0, 12).join(', '), message: `非表示列が ${s.hiddenCols.length} 列あります。`, suggestion: '意図した非表示か確認' }); });
    case 'external-link': if (model.hasExternalLinks) push({ location: 'ブック全体', message: '外部リンクのOOXMLパーツが含まれています。', suggestion: 'リンク先と必要性を確認' }); return;
    case 'formula-value-outlier': return formulaValueOutliers(model, push);
    case 'formula-pattern-outlier': return formulaPatternOutliers(model, push);
    case 'self-reference': return selfReferences(model, push);
    case 'missing-sheet-reference': return missingSheetReferences(model, push);
  }
}

function contentItems(model) {
  if (model.kind === 'word') {
    return model.paragraphs.map(p => ({
      text: p.text,
      location: paragraphLocation(p),
      context: (start, length) => excerpt(p.text, start, length)
    }));
  }
  const items = [];
  for (const sheet of model.sheets) {
    for (const cell of sheet.cells) {
      if (cell.value === null || cell.value === undefined || cell.value === '') continue;
      const text = String(cell.value);
      items.push({
        text,
        location: `シート「${sheet.name}」 ${cell.ref}`,
        context: () => text
      });
    }
  }
  return items;
}

function contentText(model) {
  return contentItems(model).map(item => item.text).join('\n');
}

function scanContentMatch(model, regex, push, message, suggestion) {
  for (const item of contentItems(model)) {
    regex.lastIndex = 0;
    let m;
    while ((m = regex.exec(item.text))) {
      push({
        location: item.location,
        matched: visibleMatch(m[0]),
        message,
        suggestion,
        context: item.context(m.index, m[0].length)
      });
      if (!regex.global) break;
    }
  }
}

function literalReplacements(rule, model, push) {
  for (const replacement of rule.replacements || []) {
    const regex = new RegExp(escapeRegExp(replacement.from), 'g');
    scanContentMatch(
      model,
      regex,
      push,
      `「${replacement.from}」の表記を確認してください。`,
      replacement.to
    );
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

function approximateRedundancy(model, push) {
  const re = /約\s*([0-9０-９][0-9０-９,，.．]*)\s*(名|人|件|個|台|回|％|%|年|か月|ヶ月|日|時間|分|円)?\s*くらい/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      push({
        location: item.location,
        matched: m[0],
        message: '「約」と「くらい」が重複しています。',
        suggestion: `約${m[1]}${m[2] || ''}`,
        context: item.context(m.index, m[0].length)
      });
    }
  }
}

function eachRedundancy(model, push) {
  const re = /各([^\s、。，．]{1,12})ごとに/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      push({
        location: item.location,
        matched: m[0],
        message: '「各」と「ごとに」の意味が重複している可能性があります。',
        suggestion: `「各${m[1]}」または「${m[1]}ごとに」に整理`,
        context: item.context(m.index, m[0].length)
      });
    }
  }
}

function startTimeYori(model, push) {
  const re = /([0-9０-９]{1,2})時より(開始|実施|開催)/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      push({
        location: item.location,
        matched: m[0],
        message: '開始点を示す「より」は「から」に言い換えられる場合があります。',
        suggestion: `${m[1]}時から${m[2]}`,
        context: item.context(m.index, m[0].length)
      });
    }
  }
}

function notationMix(model, push, a, b, label) {
  const text = contentText(model);
  const countA = countLiteral(text, a);
  const countB = countLiteral(text, b);
  if (!countA || !countB) return;
  const preferred = countA >= countB ? a : b;
  push({
    location: model.kind === 'word' ? '文書全体' : 'ブック全体',
    matched: `${a}:${countA} / ${b}:${countB}`,
    message: `${label}が「${a}」「${b}」で混在しています。`,
    suggestion: `文書方針に応じて統一（多数派: ${preferred}）`
  });
}

function yearFormatMix(model, push) {
  const text = contentText(model);
  const era = (text.match(/令和(?:元|[0-9０-９]+)年/g) || []).length;
  const western = (text.match(/(?:19|20)\d{2}年/g) || []).length;
  if (!era || !western) return;
  push({
    location: '文書全体',
    matched: `和暦 ${era} / 西暦 ${western}`,
    message: '和暦と西暦が混在しています。併記や引用を含む可能性があるため、文書の表記方針を確認してください。',
    suggestion: '意図した併記を除き、年表記を統一'
  });
}

function tonUnitMix(model, push) {
  const text = contentText(model);
  const ton = (text.match(/[0-9０-９][0-9０-９,，.．]*\s*トン/g) || []).length;
  const t = (text.match(/[0-9０-９][0-9０-９,，.．]*\s*t(?![A-Za-z])/g) || []).length;
  if (!ton || !t) return;
  push({
    location: model.kind === 'word' ? '文書全体' : 'ブック全体',
    matched: `トン:${ton} / t:${t}`,
    message: '数値に続く質量単位「トン」と「t」が混在しています。',
    suggestion: '文書方針に応じて統一'
  });
}

function nthDay(model, push) {
  const re = /第([0-9０-９]+)日目/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      push({
        location: item.location,
        matched: m[0],
        message: '「第」と「目」が重複的な序数表現になっていないか確認してください。',
        suggestion: `第${m[1]}日`,
        context: item.context(m.index, m[0].length)
      });
    }
  }
}

function fullwidthUnit(model, push) {
  const re = /(?:ｋＷ|ＫＷ|ｋｗ|ｃｍ|ＣＭ|ｍｍ|ＭＭ|ｐｔ|ＰＴ|ｋｇ|ＫＧ|ＭＷ)/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      push({
        location: item.location,
        matched: m[0],
        message: '単位に全角英字が使われています。',
        suggestion: m[0].normalize('NFKC'),
        context: item.context(m.index, m[0].length)
      });
    }
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

function captionDuplicates(model, push) {
  const seen = new Map();
  for (const p of model.paragraphs) {
    const m = p.text.match(/^\s*(図|表)\s*([0-9０-９]+(?:\s*[-－―ー]\s*[0-9０-９]+)+)/);
    if (!m) continue;
    const key = normalizeFigureKey(m[1], m[2]);
    if (seen.has(key)) {
      push({
        location: paragraphLocation(p),
        matched: m[0].trim(),
        message: `同じ図表番号が既に ${paragraphLocation(seen.get(key))} で使われています。`,
        suggestion: '図表番号の重複を確認'
      });
    } else {
      seen.set(key, p);
    }
  }
}

function headingEdgeSpaces(model, push) {
  for (const p of model.paragraphs) {
    if (!isHeading(p) || !p.text) continue;
    const leading = p.text.match(/^[ \u3000]+/);
    const trailing = p.text.match(/[ \u3000]+$/);
    if (leading) push({
      location: paragraphLocation(p),
      matched: visibleSpaces(leading[0]),
      message: '見出し先頭に空白があります。',
      suggestion: '不要であれば削除',
      context: excerpt(p.text, 0, leading[0].length)
    });
    if (trailing) push({
      location: paragraphLocation(p),
      matched: visibleSpaces(trailing[0]),
      message: '見出し末尾に空白があります。',
      suggestion: '削除',
      context: excerpt(p.text, p.text.length - trailing[0].length, trailing[0].length)
    });
  }
}

function paragraphTrailingSpaces(model, push) {
  for (const p of model.paragraphs) {
    if (!p.text || isHeading(p)) continue;
    const m = p.text.match(/[ \u3000]+$/);
    if (!m) continue;
    push({
      location: paragraphLocation(p),
      matched: visibleSpaces(m[0]),
      message: '段落末尾に空白があります。',
      suggestion: '不要であれば削除',
      context: excerpt(p.text, p.text.length - m[0].length, m[0].length)
    });
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

function missingSheetReferences(model, push) {
  const names = new Set(model.sheets.map(s => s.name));
  const re = /'((?:[^']|'')+)'!|([\p{L}\p{N}_][\p{L}\p{N}_.]*)!/gu;
  eachCell(model, ({ sheet, cell }) => {
    if (!cell.formula || (cell.formula.includes('[') && cell.formula.includes(']'))) return;
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(cell.formula))) {
      const rawName = (m[1] || m[2] || '').replace(/''/g, "'");
      if (!rawName || names.has(rawName)) continue;
      push(cellFinding(
        sheet,
        cell,
        `数式が存在しないシート「${rawName}」を参照している可能性があります。`,
        'シート名変更・削除・参照式を確認',
        rawName
      ));
    }
  });
}

function eachCell(model, fn) { for (const sheet of model.sheets) for (const cell of sheet.cells) fn({ sheet, cell }); }
function cellFinding(sheet, cell, message, suggestion, matched = null) { return { location: `シート「${sheet.name}」 ${cell.ref}`, matched: matched ?? String(cell.value || cell.formula || '—'), message, suggestion, context: cell.formula ? `=${cell.formula}` : String(cell.value ?? '') }; }
function paragraphLocation(p) { return `${p.pageHint ? `保存時ページ目安 ${p.pageHint} / ` : ''}第${p.index}段落`; }
function visibleMatch(s) { return s === '\u00a0' ? 'U+00A0' : s.replace(/\u00a0/g, '⍽'); }
function visibleSpaces(s) { return Array.from(s).map(ch => ch === '\u3000' ? '全角空白' : '半角空白').join(' + '); }
function excerpt(text, start, length) { const a = Math.max(0, start - 24), b = Math.min(text.length, start + length + 36); return `${a ? '…' : ''}${text.slice(a, b)}${b < text.length ? '…' : ''}`; }
function toAsciiDigits(s) { return s.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)); }
function normalizeFigureKey(kind, num) { return `${kind}:${toAsciiDigits(num).replace(/\s/g, '').replace(/[－―ー]/g, '-')}`; }
function summarizeNumbers(nums) { const shown = nums.slice(0, 15).join(', '); return nums.length > 15 ? `${shown} …` : shown; }
function columnName(num) { let s = ''; for (let n = num; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; }
function parseCellRef(ref) { const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref); if (!m) return { row: 0, col: 0 }; let col = 0; for (const ch of m[1].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64; return { row: Number(m[2]), col }; }
function key(row, col) { return `${row}:${col}`; }
function cellMap(cells) { const m = new Map(); for (const c of cells) { const p = parseCellRef(c.ref); if (p.row && p.col) m.set(key(p.row, p.col), c); } return m; }
function hasFormula(map, row, col) { return Boolean(row > 0 && col > 0 && map.get(key(row, col))?.formula); }
function escapeRegExp(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function countLiteral(text, token) { return text.split(token).length - 1; }
function isHeading(p) { const id = String(p.styleId || '').toLowerCase(); return id.includes('heading') || id.includes('見出し'); }

function normalizeFormula(formula, baseRow, baseCol) {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/gi, (full, absCol, letters, absRow, rowText) => {
    const refCol = parseCellRef(`${letters}${rowText}`).col;
    const refRow = Number(rowText);
    const c = absCol ? `C$${refCol}` : `C[${refCol - baseCol}]`;
    const r = absRow ? `R$${refRow}` : `R[${refRow - baseRow}]`;
    return `${r}${c}`;
  });
}
