import { RULES, CATEGORY_LABELS, runDocumentQc, groupFindingsByRule } from '../core.js';
import { readCurrentWordDocument, getOfficeEnvironmentLabel } from './document-reader.js';
import { locateFinding, applyFindingFix, applyFindingFixes } from './word-actions.js';

const state = {
  model: null,
  filename: '',
  source: '',
  findings: [],
  enabledRuleIds: new Set(RULES.map(rule => rule.id)),
  busy: false
};

const $ = id => document.getElementById(id);

Office.onReady(info => {
  if (info.host && info.host !== Office.HostType.Word) {
    showError('Document QC は Word 用のアドインです。');
    $('checkButton').disabled = true;
    return;
  }
  init();
});

function init() {
  $('environmentBadge').textContent = getOfficeEnvironmentLabel();
  $('checkButton').addEventListener('click', executeCheck);
  $('resultSearch').addEventListener('input', renderFindings);
  $('severityFilter').addEventListener('change', renderFindings);
  document.querySelectorAll('.summary-stat').forEach(button => button.addEventListener('click', () => {
    const select = $('severityFilter');
    select.value = select.value === button.dataset.severity ? 'all' : button.dataset.severity;
    renderFindings();
  }));
  $('resetRules').addEventListener('click', () => {
    state.enabledRuleIds = new Set(RULES.map(rule => rule.id));
    renderRules();
    recalculateCurrentModel();
  });
  renderRules();
  updateEnabledRuleCount();
}

async function executeCheck() {
  if (state.busy) return;
  setBusy(true);
  hideError();
  try {
    await performCheck();
  } catch (error) {
    console.error(error);
    showError(error instanceof Error ? error.message : String(error));
  } finally {
    setBusy(false);
  }
}

async function performCheck() {
  $('progress').hidden = false;
  setProgress(2, '開始しています');
  const { model, filename, source } = await readCurrentWordDocument(setProgress);
  state.model = model;
  state.filename = filename;
  state.source = source;
  state.findings = runDocumentQc(model, state.enabledRuleIds, setProgress);
  setProgress(100, `完了: ${state.findings.length}件`);
  renderResults();
}

function recalculateCurrentModel() {
  if (!state.model) return;
  state.findings = runDocumentQc(state.model, state.enabledRuleIds);
  renderResults();
}

function renderResults() {
  $('summarySection').hidden = false;
  const counts = countBySeverity(state.findings);
  $('errorCount').textContent = String(counts.error);
  $('confirmCount').textContent = String(counts.confirm);
  $('infoCount').textContent = String(counts.info);
  $('resultMeta').textContent = `${state.filename || '現在の文書'} / ${state.findings.length}件 / ${state.enabledRuleIds.size}ルール有効`;
  renderFindings();
}

function renderFindings() {
  const list = $('resultsList');
  list.textContent = '';
  if (!state.model) return;

  const q = $('resultSearch').value.trim().toLowerCase();
  const severity = $('severityFilter').value;
  const filtered = state.findings.filter(finding => {
    if (severity !== 'all' && finding.severity !== severity) return false;
    if (!q) return true;
    return `${finding.ruleId} ${finding.title} ${finding.location} ${finding.matched} ${finding.message} ${finding.suggestion} ${finding.context}`.toLowerCase().includes(q);
  });

  $('cleanState').hidden = state.findings.length !== 0;
  $('filteredEmpty').hidden = state.findings.length === 0 || filtered.length !== 0;
  if (!filtered.length) return;

  for (const group of groupFindingsByRule(filtered)) list.append(renderFindingGroup(group));
}

function renderFindingGroup(group) {
  const root = el('section', 'finding-group');
  const header = el('div', 'finding-group-header');
  const title = el('div', 'rule-title');
  title.append(el('code', '', group.ruleId), el('strong', '', group.title || group.ruleId));
  header.append(title, el('span', 'count-pill', `${group.findings.length}件`));

  const safeFixes = group.findings.filter(finding => finding.fix?.mode === 'safe');
  if (safeFixes.length > 1 && safeFixes.length === group.findings.length) {
    const bulk = el('button', 'button bulk-button', `${safeFixes.length}件を修正`);
    bulk.addEventListener('click', () => withFindingAction(async () => {
      await applyFindingFixes(safeFixes);
      await performCheck();
    }));
    header.append(bulk);
  }
  root.append(header);

  for (const finding of group.findings) root.append(renderFinding(finding));
  return root;
}

function renderFinding(finding) {
  const root = el('article', 'finding-item');
  const top = el('div', 'finding-topline');
  top.append(severityBadge(finding.severity), el('span', 'location', finding.location || ''));
  root.append(top);

  if (finding.matched && finding.matched !== '—') root.append(el('div', 'match-box', finding.matched));
  root.append(el('p', 'finding-message', finding.message || finding.title || ''));

  if (finding.suggestion && finding.suggestion !== '—') {
    const suggestion = el('p', 'suggestion');
    suggestion.append(el('strong', '', '候補: '), document.createTextNode(finding.suggestion));
    root.append(suggestion);
  }
  if (finding.context) root.append(el('p', 'context', finding.context));

  const actions = el('div', 'finding-actions');
  if (finding.anchor) {
    const locate = el('button', 'button small', '移動');
    locate.addEventListener('click', () => withFindingAction(() => locateFinding(finding)));
    actions.append(locate);
  }
  if (finding.fix) {
    const label = finding.fix.mode === 'safe' ? '修正' : '候補を適用';
    const fix = el('button', `button small${finding.fix.mode === 'safe' ? ' primary' : ''}`, label);
    fix.addEventListener('click', () => withFindingAction(async () => {
      await applyFindingFix(finding);
      await performCheck();
    }));
    actions.append(fix);
  }
  if (actions.childNodes.length) root.append(actions);
  return root;
}

function renderRules() {
  const container = $('rulesList');
  container.textContent = '';
  for (const [category, label] of Object.entries(CATEGORY_LABELS)) {
    const rules = RULES.filter(rule => rule.category === category && rule.appliesTo.includes('word'));
    if (!rules.length) continue;
    const group = el('section', 'rule-group');
    group.append(el('h3', '', `${label} (${rules.length})`));
    for (const rule of rules) {
      const row = el('label', 'rule-row');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = state.enabledRuleIds.has(rule.id);
      checkbox.addEventListener('change', () => {
        checkbox.checked ? state.enabledRuleIds.add(rule.id) : state.enabledRuleIds.delete(rule.id);
        updateEnabledRuleCount();
        recalculateCurrentModel();
      });
      row.append(checkbox, el('code', '', rule.id), el('span', '', rule.title));
      group.append(row);
    }
    container.append(group);
  }
  updateEnabledRuleCount();
}

async function withFindingAction(action) {
  if (state.busy) return;
  setBusy(true);
  hideError();
  try {
    await action();
  } catch (error) {
    console.error(error);
    showError(error instanceof Error ? error.message : String(error));
  } finally {
    setBusy(false);
  }
}

function setBusy(value) {
  state.busy = value;
  $('checkButton').disabled = value;
  document.querySelectorAll('.finding-actions button, .bulk-button').forEach(button => { button.disabled = value; });
}

function setProgress(percent, text) {
  $('progressBar').style.width = `${Math.max(0, Math.min(100, percent))}%`;
  $('progressText').textContent = text;
}

function showError(message) {
  $('errorBox').hidden = false;
  $('errorBox').textContent = message;
}

function hideError() {
  $('errorBox').hidden = true;
  $('errorBox').textContent = '';
}

function updateEnabledRuleCount() {
  const wordTotal = RULES.filter(rule => rule.appliesTo.includes('word')).length;
  const wordEnabled = RULES.filter(rule => rule.appliesTo.includes('word') && state.enabledRuleIds.has(rule.id)).length;
  $('enabledRuleCount').textContent = `${wordEnabled}/${wordTotal}`;
}

function countBySeverity(findings) {
  const counts = { error: 0, confirm: 0, info: 0 };
  for (const finding of findings) counts[finding.severity] = (counts[finding.severity] || 0) + 1;
  return counts;
}

function severityBadge(severity) {
  const label = severity === 'error' ? '要修正' : severity === 'confirm' ? '確認' : '参考';
  return el('span', `severity ${severity}`, label);
}

function el(tag, className = '', text = '') {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== '') node.textContent = text;
  return node;
}
