import assert from 'node:assert/strict';
import { runRules } from '../src/engine.js';

const word = {
  kind: 'word',
  paragraphs: [
    {
      index: 1,
      pageHint: 1,
      styleId: 'Heading1',
      text: '　発砲スチロール SDGS 約20名くらい ',
      runs: []
    },
    {
      index: 2,
      pageHint: 1,
      styleId: 'Normal',
      text: '工法を比較し、嫌気性消化の条件を整理した。',
      runs: []
    }
  ],
  tables: [],
  hasComments: false,
  hasTrackedChanges: false,
  text: '　発砲スチロール SDGS 約20名くらい \n工法を比較し、嫌気性消化の条件を整理した。'
};

const wordFindings = runRules(word);
const wordIds = new Set(wordFindings.map(f => f.ruleId));
assert(wordIds.has('TXT-006'), 'obvious typo should be detected');
assert(wordIds.has('TXT-008'), 'SDGs casing should be detected');
assert(wordIds.has('TXT-025'), 'approximate redundancy should be detected');
assert(wordIds.has('DOC-009'), 'heading edge space should be detected');
assert(!wordFindings.some(f => f.matched === '工法'), 'valid word 工法 must not be auto-corrected');
assert(!wordFindings.some(f => f.matched === '消化'), 'valid technical term 消化 must not be auto-corrected');

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

console.log(`rule tests passed: Word ${wordFindings.length}, Excel ${excelFindings.length}`);
