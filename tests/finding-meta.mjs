import assert from 'node:assert/strict';
import { decorateFindings, groupFindingsByRule } from '../src/finding-meta.js';

const model = {
  kind: 'word',
  paragraphs: [
    { index: 1, text: 'A\u00a0B A\u00a0B' },
    { index: 2, text: '発砲スチロールと第 6 章　' }
  ]
};

const findings = decorateFindings(model, [
  { ruleId: 'TXT-001', title: 'NBSP', severity: 'error', location: '第1段落', matched: 'U+00A0', suggestion: '通常の空白または不要なら削除' },
  { ruleId: 'TXT-001', title: 'NBSP', severity: 'error', location: '第1段落', matched: 'U+00A0', suggestion: '通常の空白または不要なら削除' },
  { ruleId: 'TXT-006', title: '誤記', severity: 'error', location: '第2段落', matched: '発砲スチロール', suggestion: '発泡スチロール' },
  { ruleId: 'TXT-046', title: '1桁', severity: 'confirm', location: '第2段落', matched: '6', suggestion: '６' },
  { ruleId: 'DOC-010', title: '末尾空白', severity: 'confirm', location: '第2段落', matched: '全角空白', suggestion: '不要であれば削除' },
  { ruleId: 'TXT-005', title: '文体混在', severity: 'confirm', location: '文書全体', matched: '敬体 1 / 常体 2', suggestion: '文書方針に合わせて確認' }
]);

assert.equal(findings[0].anchor.searchText, '\u00a0');
assert.equal(findings[0].anchor.occurrence, 0);
assert.equal(findings[1].anchor.occurrence, 1);
assert.deepEqual(findings[0].fix, { mode: 'safe', replacement: ' ' });
assert.deepEqual(findings[2].fix, { mode: 'safe', replacement: '発泡スチロール' });
assert.deepEqual(findings[3].fix, { mode: 'review', replacement: '６' });
assert.equal(findings[4].anchor.searchText, '\u3000');
assert.deepEqual(findings[4].fix, { mode: 'review', replacement: '' });
assert.equal(findings[5].anchor, null);
assert.equal(findings[5].fix, null);

const groups = groupFindingsByRule(findings);
assert.equal(groups.length, 5);
assert.equal(groups.find(group => group.ruleId === 'TXT-001').findings.length, 2);

console.log('finding metadata tests passed');
