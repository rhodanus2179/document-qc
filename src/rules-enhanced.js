import { RULES as BASE_RULES, CATEGORY_LABELS as BASE_CATEGORY_LABELS } from './rules.js';

export const CATEGORY_LABELS = {
  ...BASE_CATEGORY_LABELS,
  reference: '公開情報照合'
};

export const EXTRA_RULES = [
  {
    id: 'TXT-041', category: 'text', severity: 'confirm', engine: 'source-typo-context',
    title: '出典文脈の「出展」',
    description: '資料・図表・文献などの出所を示す文脈で「出展」が使われている場合、「出典」の誤記候補として確認します。',
    appliesTo: ['word']
  },
  {
    id: 'TXT-042', category: 'text', severity: 'confirm', engine: 'pruning-branch-context',
    title: '処理文脈の「選定枝」',
    description: '造園・廃棄物処理等の文脈で「選定枝」が使われている場合、「剪定枝」の誤記候補として確認します。',
    appliesTo: ['word', 'excel']
  },
  {
    id: 'TXT-043', category: 'text', severity: 'confirm', engine: 'discrepancy-typo-context',
    title: '食い違い文脈の「祖語」',
    description: '食い違い・不一致を意味する文脈で「祖語」が使われている場合、「齟齬」の誤記候補として確認します。',
    appliesTo: ['word']
  },
  {
    id: 'TXT-044', category: 'text', severity: 'confirm', engine: 'literal-extra',
    title: '「要素の感覚」',
    description: 'グラフ・レイアウト等で「要素の間隔」を意図していないか確認します。',
    appliesTo: ['word', 'excel'],
    replacements: [{ from: '要素の感覚', to: '要素の間隔' }]
  },
  {
    id: 'TXT-045', category: 'text', severity: 'confirm', engine: 'literal-extra',
    title: '公用文調の「または」',
    description: '法令・通知等に近い公用文調で「又は」を採用する方針の場合、「または」を確認します。一般向け文章では必ずしも修正不要です。',
    appliesTo: ['word'],
    replacements: [{ from: 'または', to: '又は' }]
  },
  {
    id: 'REF-001', category: 'reference', severity: 'confirm', engine: 'public-official-name',
    title: '第六次環境基本計画の正式名称',
    description: '環境省公開情報に基づき、「第六次環境基本計画」の正式名称表記を確認します。',
    appliesTo: ['word', 'excel']
  },
  {
    id: 'REF-002', category: 'reference', severity: 'confirm', engine: 'public-decision-date',
    title: '第六次環境基本計画の閣議決定年',
    description: '環境省公開情報（令和6年5月21日閣議決定）と、同一箇所に記載された年を照合します。',
    appliesTo: ['word', 'excel']
  }
];

export const RULES = [...BASE_RULES, ...EXTRA_RULES];
