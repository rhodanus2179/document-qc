const SAFE_FIX_RULES = new Set([
  'TXT-001', // NBSP -> normal space
  'TXT-006', // 発砲スチロール -> 発泡スチロール
  'TXT-020', // ストーカー式 -> ストーカ式
  'TXT-042', // 選定枝 -> 剪定枝 (context-limited high confidence)
  'TXT-044'  // 要素の感覚 -> 要素の間隔
]);

const REVIEW_FIX_RULES = new Set([
  'TXT-002', 'TXT-003',
  'TXT-007', 'TXT-008', 'TXT-009', 'TXT-010', 'TXT-011', 'TXT-012', 'TXT-013', 'TXT-014',
  'TXT-015', 'TXT-016', 'TXT-017', 'TXT-018', 'TXT-019', 'TXT-021', 'TXT-022', 'TXT-023',
  'TXT-024', 'TXT-025', 'TXT-027', 'TXT-028', 'TXT-029', 'TXT-031', 'TXT-038', 'TXT-040',
  'TXT-041', 'TXT-043', 'TXT-045', 'TXT-046',
  'REF-001', 'REF-002',
  'DOC-009', 'DOC-010'
]);

const GENERIC_SUGGESTIONS = [
  '確認', '文書方針', '意図', '統一（', '必要性', '参照先', '図表番号', '欠番', '重複', '校正跡', '変更履歴'
];

/**
 * Add stable, UI-oriented metadata to deterministic engine findings without
 * coupling the rule engine itself to Word/Office.js.
 */
export function decorateFindings(model, findings) {
  const cursors = new Map();
  return findings.map((finding, index) => {
    const anchor = model?.kind === 'word' ? buildWordAnchor(model, finding, cursors) : null;
    const replacement = anchor ? inferReplacement(finding, anchor.searchText) : null;
    const mode = replacement === null ? 'none'
      : SAFE_FIX_RULES.has(finding.ruleId) ? 'safe'
      : REVIEW_FIX_RULES.has(finding.ruleId) ? 'review'
      : 'none';

    return {
      ...finding,
      findingId: makeFindingId(finding, anchor, index),
      anchor,
      fix: mode === 'none' ? null : { mode, replacement }
    };
  });
}

export function groupFindingsByRule(findings) {
  const groups = new Map();
  for (const finding of findings) {
    if (!groups.has(finding.ruleId)) {
      groups.set(finding.ruleId, {
        ruleId: finding.ruleId,
        title: finding.title,
        severity: finding.severity,
        findings: []
      });
    }
    const group = groups.get(finding.ruleId);
    group.findings.push(finding);
    if (severityRank(finding.severity) > severityRank(group.severity)) group.severity = finding.severity;
  }
  return Array.from(groups.values());
}

function buildWordAnchor(model, finding, cursors) {
  const paragraphIndex = parseParagraphIndex(finding.location);
  if (!paragraphIndex) return null;
  const paragraph = model.paragraphs?.[paragraphIndex - 1];
  if (!paragraph) return null;

  const needle = inferNeedle(finding);
  if (!needle) return null;

  const key = `${finding.ruleId}\u0000${paragraphIndex}\u0000${needle}`;
  const from = cursors.get(key) ?? 0;
  let start = paragraph.text.indexOf(needle, from);
  if (start < 0 && from > 0) start = paragraph.text.indexOf(needle);
  if (start < 0) return null;
  cursors.set(key, start + Math.max(needle.length, 1));

  const occurrence = countBefore(model.paragraphs, paragraphIndex - 1, start, needle);
  return {
    kind: 'word-text',
    paragraphIndex,
    start,
    length: needle.length,
    searchText: needle,
    occurrence,
    paragraphText: paragraph.text
  };
}

function inferNeedle(finding) {
  if (finding.ruleId === 'TXT-001') return '\u00a0';
  if (finding.ruleId === 'DOC-009' || finding.ruleId === 'DOC-010') {
    return decodeVisibleSpaces(String(finding.matched || ''));
  }

  const matched = String(finding.matched ?? '');
  if (!matched || matched === '—' || matched === '(空文字)') return null;
  if (/^(?:敬体|和暦|トン:|％:|CO2:|／:)/.test(matched)) return null;
  if (/^[^\s]+:\d+\s*\/\s*[^\s]+:\d+/.test(matched)) return null;
  return matched.includes('⍽') ? matched.replace(/⍽/g, '\u00a0') : matched;
}

function inferReplacement(finding, needle) {
  if (!needle) return null;
  switch (finding.ruleId) {
    case 'TXT-001': return ' ';
    case 'TXT-002': return ' ';
    case 'TXT-003': return firstUnicodeChar(needle);
    case 'DOC-009':
    case 'DOC-010': return '';
  }

  const suggestion = String(finding.suggestion ?? '').trim();
  if (!suggestion || suggestion === '—') return null;
  if (suggestion.includes('または') || suggestion.includes('又は')) {
    // TXT-045 has a literal replacement, but alternatives in other rules must stay manual.
    if (finding.ruleId !== 'TXT-045') return null;
  }
  if (GENERIC_SUGGESTIONS.some(token => suggestion.includes(token))) return null;
  if (/^削除$/.test(suggestion)) return '';
  return suggestion;
}

function countBefore(paragraphs, paragraphOffset, start, needle) {
  let count = 0;
  for (let i = 0; i < paragraphOffset; i++) count += countNonOverlapping(paragraphs[i]?.text || '', needle);
  count += countNonOverlapping((paragraphs[paragraphOffset]?.text || '').slice(0, start), needle);
  return count;
}

function countNonOverlapping(text, needle) {
  if (!needle) return 0;
  let count = 0;
  let from = 0;
  while (from <= text.length - needle.length) {
    const hit = text.indexOf(needle, from);
    if (hit < 0) break;
    count++;
    from = hit + Math.max(needle.length, 1);
  }
  return count;
}

function parseParagraphIndex(location) {
  const match = /第(\d+)段落/.exec(String(location || ''));
  return match ? Number(match[1]) : 0;
}

function decodeVisibleSpaces(value) {
  if (!value) return '';
  const parts = value.split(/\s*\+\s*/);
  if (!parts.every(part => part === '半角空白' || part === '全角空白')) return '';
  return parts.map(part => part === '全角空白' ? '\u3000' : ' ').join('');
}

function firstUnicodeChar(value) {
  return Array.from(value)[0] || '';
}

function makeFindingId(finding, anchor, index) {
  if (anchor) return `${finding.ruleId}:${anchor.paragraphIndex}:${anchor.start}:${index}`;
  return `${finding.ruleId}:global:${index}`;
}

function severityRank(value) {
  return value === 'error' ? 3 : value === 'confirm' ? 2 : 1;
}
