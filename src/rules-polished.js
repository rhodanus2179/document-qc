import { RULES as ENHANCED_RULES, CATEGORY_LABELS } from './rules-enhanced.js';

const LEGACY_SINGLE_DIGIT_RULE = {
  id: 'TXT-046',
  category: 'text',
  severity: 'confirm',
  engine: 'legacy-single-digit-width',
  title: '旧来表記：1桁数字は全角',
  description: '古い公文書・ワープロ系の表記慣行として、独立した1桁の半角数字を全角にする方針の場合に確認します。現在の一般的な文書で必須の規則ではありません。',
  appliesTo: ['word']
};

export { CATEGORY_LABELS };
export const EXTRA_POLISH_RULES = [LEGACY_SINGLE_DIGIT_RULE];

export const RULES = [
  ...ENHANCED_RULES.map(rule => {
    if (rule.id === 'TXT-042' || rule.id === 'TXT-044') return { ...rule, severity: 'error' };
    return rule;
  }),
  ...EXTRA_POLISH_RULES
];
