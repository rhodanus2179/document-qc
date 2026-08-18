import { runRules as runEnhancedRules } from './engine-enhanced.js';
import { RULES, EXTRA_POLISH_RULES } from './rules-polished.js';

export function runRules(model, enabledIds = new Set(RULES.map(rule => rule.id)), onProgress = () => {}) {
  const findings = runEnhancedRules(model, enabledIds, onProgress).map(polishFinding);

  const applicable = EXTRA_POLISH_RULES.filter(rule => enabledIds.has(rule.id) && rule.appliesTo.includes(model.kind));
  for (const rule of applicable) {
    const push = item => findings.push({
      ruleId: rule.id,
      category: rule.category,
      severity: rule.severity,
      title: rule.title,
      matched: '—',
      suggestion: '—',
      context: '',
      ...item
    });

    if (rule.engine === 'legacy-single-digit-width') legacySingleDigitWidth(model, push);
  }

  return findings;
}

function polishFinding(finding) {
  if (finding.ruleId === 'TXT-042' || finding.ruleId === 'TXT-044') {
    return { ...finding, severity: 'error' };
  }

  if (finding.ruleId === 'REF-002') {
    const matched = String(finding.matched || '');
    const suggestion = matched.includes('令和') ? '令和6年' : /^\s*(?:19|20)/.test(matched) ? '2024年' : finding.suggestion;
    return { ...finding, suggestion };
  }

  return finding;
}

function legacySingleDigitWidth(model, push) {
  if (model.kind !== 'word') return;

  for (const paragraph of model.paragraphs) {
    const text = paragraph.text || '';
    for (let index = 0; index < text.length; index++) {
      const ch = text[index];
      if (!/[0-9]/.test(ch)) continue;
      const prev = text[index - 1] || '';
      const next = text[index + 1] || '';

      // Do not split multi-digit values, decimals, comma-grouped numbers, or identifiers such as A1.
      if (/[A-Za-z0-9０-９,，.．]/.test(prev) || /[A-Za-z0-9０-９,，.．]/.test(next)) continue;

      push({
        location: paragraphLocation(paragraph),
        matched: ch,
        message: '旧来の公文書・ワープロ系の表記慣行では、1桁の数字を全角とする場合があります。現在の一般的な文書で必須の規則ではないため、文書方針に応じて確認してください。',
        suggestion: toFullwidthDigit(ch),
        context: excerpt(text, index, 1)
      });
    }
  }
}

function toFullwidthDigit(digit) {
  return String.fromCharCode(digit.charCodeAt(0) + 0xFEE0);
}

function paragraphLocation(paragraph) {
  return `${paragraph.pageHint ? `保存時ページ目安 ${paragraph.pageHint} / ` : ''}第${paragraph.index}段落`;
}

function excerpt(text, start, length) {
  const a = Math.max(0, start - 24);
  const b = Math.min(text.length, start + length + 36);
  return `${a ? '…' : ''}${text.slice(a, b)}${b < text.length ? '…' : ''}`;
}
