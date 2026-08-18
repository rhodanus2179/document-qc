import assert from 'node:assert/strict';
import { runRules } from '../src/engine-polished.js';

function wordModel(paragraphTexts) {
  const paragraphs = paragraphTexts.map((text, index) => ({
    index: index + 1,
    pageHint: 1,
    styleId: index === 0 ? 'Heading1' : 'Normal',
    text,
    runs: []
  }));
  return {
    kind: 'word',
    paragraphs,
    tables: [],
    hasComments: false,
    hasTrackedChanges: false,
    text: paragraphs.map(p => p.text).join('\n')
  };
}

const word = wordModel([
  '　発砲スチロール SDGS 約20名くらい ',
  '工法を比較し、嫌気性消化の条件を整理した。'
]);

const wordFindings = runRules(word);
const wordIds = new Set(wordFindings.map(f => f.ruleId));
assert(wordIds.has('TXT-006'), 'obvious typo should be detected');
assert(wordIds.has('TXT-008'), 'SDGs casing should be detected');
assert(wordIds.has('TXT-025'), 'approximate redundancy should be detected');
assert(wordIds.has('DOC-009'), 'heading edge space should be detected');
assert(!wordFindings.some(f => f.matched === '工法'), 'valid word 工法 must not be auto-corrected');
assert(!wordFindings.some(f => f.matched === '消化'), 'valid technical term 消化 must not be auto-corrected');

const styleWord = wordModel([
  'この文書は動作確認用です。わざと不備を混ぜてあります。',
  '本業務では、施設の整備方針を検討した。調査の結果は次のとおりである。',
  '調査対象は次のとおりです。',
  '回収率は80％であった。詳細は別紙である。',
  'いずれかを選択する方針とする。'
]);
const styleFinding = runRules(styleWord).find(f => f.ruleId === 'TXT-005');
assert(styleFinding, 'register mixing should be detected');
assert.equal(styleFinding.matched, '敬体 3 / 常体 4', 'register mixing must count sentence endings, not paragraphs');

const parityWord = wordModel([
  '資料の出展は各章の末尾に記載する。展示会に出展する。出展資料を配布した。',
  '発砲スチロールと選定枝の処理方法を整理した。',
  '評価にあたっては、担当者間で祖語が生じないよう基準を共有した。日本祖語を研究する。',
  'グラフでは要素の感覚を調整している。',
  '現行方式または新方式のいずれかを選択する。',
  '第 6 次環境基本計画（令和 5 年閣議決定）の内容を踏まえている。'
]);
const parityFindings = runRules(parityWord);
const parityIds = new Set(parityFindings.map(f => f.ruleId));
for (const id of ['TXT-041', 'TXT-042', 'TXT-043', 'TXT-044', 'TXT-045', 'TXT-046', 'REF-001', 'REF-002']) {
  assert(parityIds.has(id), `${id} should be detected in the demo-style document`);
}
assert.equal(parityFindings.filter(f => f.ruleId === 'TXT-041').length, 1, 'legitimate 出展 / 出展資料 must not be flagged');
assert.equal(parityFindings.filter(f => f.ruleId === 'TXT-043').length, 1, 'legitimate 日本祖語 must not be flagged');
assert.equal(parityFindings.find(f => f.ruleId === 'TXT-042')?.severity, 'error', 'high-confidence 選定枝 finding should be an error');
assert.equal(parityFindings.find(f => f.ruleId === 'TXT-044')?.severity, 'error', '要素の感覚 should be an error');
assert.equal(parityFindings.find(f => f.ruleId === 'REF-001')?.suggestion, '第六次環境基本計画');
assert.equal(parityFindings.find(f => f.ruleId === 'REF-002')?.suggestion, '令和6年');

const widthWord = wordModel([
  '第 6 章、令和 5 年。数量は1,200件、比率は0.5、識別子A1。'
]);
const widthFindings = runRules(widthWord).filter(f => f.ruleId === 'TXT-046');
assert.equal(widthFindings.length, 2, 'legacy single-digit rule must ignore grouped numbers, decimals, and identifiers');
assert.deepEqual(widthFindings.map(f => f.matched), ['6', '5']);
assert.deepEqual(widthFindings.map(f => f.suggestion), ['６', '５']);
assert(widthFindings.every(f => f.severity === 'confirm'), 'legacy digit-width rule must stay confirmation-only');

const excel = {
  kind: 'excel',
  sheets: [
    {
      name: 'Sheet1',
      state: 'visible',
      hiddenRows: [],
      hiddenCols: [],
      cells: [
        { ref: 'A1', value: 'KW 10％', formula: null },
        { ref: 'A2', value: '10%', formula: null },
        { ref: 'B1', value: '', formula: "'Missing'!A1" }
      ]
    }
  ],
  hasExternalLinks: false
};

const excelFindings = runRules(excel);
const excelIds = new Set(excelFindings.map(f => f.ruleId));
assert(excelIds.has('TXT-024'), 'kW casing should be detected in Excel cells');
assert(excelIds.has('TXT-032'), 'percent width mixing should be detected');
assert(excelIds.has('XLS-010'), 'missing sheet references should be detected');
assert(!excelIds.has('TXT-046'), 'legacy single-digit width rule is Word-only');

console.log(`rule tests passed: Word ${wordFindings.length}, parity ${parityFindings.length}, width ${widthFindings.length}, Excel ${excelFindings.length}`);
