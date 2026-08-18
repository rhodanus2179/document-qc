import { parseOfficeFile } from '../ooxml.js';
import { RULES, CATEGORY_LABELS, runDocumentQc } from '../core.js';
import { downloadCsv, downloadXlsx } from '../export.js';

const state = {
  file: null,
  model: null,
  findings: [],
  enabledRuleIds: new Set(RULES.map(rule => rule.id))
};

const $ = id => document.getElementById(id);
const fileInput = $('fileInput');
const dropZone = $('dropZone');
const checkButton = $('checkButton');
const clearButton = $('clearButton');
const selectedFile = $('selectedFile');
const progress = $('progress');
const progressBar = $('progressBar');
const progressText = $('progressText');
const errorBox = $('errorBox');

init();

function init() {
  updateVersionUi();
  renderRuleSummary();
  renderRules();
  bindNavigation();
  bindFileSelection();
  bindResults();
  updateResultsUi();
  if (!window.__DOCUMENT_QC_NETWORK_LOCKED__) showError('ネットワークロックの初期化を確認できません。');
}


function updateVersionUi() {
  const footerVersion = document.querySelector('footer span:first-child');
  if (footerVersion) footerVersion.textContent = 'Document QC v1.1.0';
  const aboutLead = document.querySelector('#view-about .prose > p');
  if (aboutLead) aboutLead.innerHTML = 'Document QC v1.1.0 は OOXML（<code>.docx</code> / <code>.xlsx</code>）をブラウゞ上で直接解析する、完全静的な品質チェックツールです。Word Add-in 版と同じ QC コアを使用します。';
}

function bindNavigation() {
  document.querySelectorAll('.tab').forEach(button => button.addEventListener('click', () => activateView(button.dataset.view)));
}

function activateView(name) {
  document.querySelectorAll('.tab').forEach(button => button.classList.toggle('is-active', button.dataset.view === name));
  document.querySelectorAll('.view').forEach(view => view.classList.toggle('is-active', view.id === `view-${name}`));
}

function bindFileSelection() {
  fileInput.addEventListener('change', () => setFile(fileInput.files?.[0] || null));
  dropZone.addEventListener('dragover', event => { event.preventDefault(); dropZone.classList.add('is-dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('is-dragover'));
  dropZone.addEventListener('drop', event => {
    event.preventDefault();
    dropZone.classList.remove('is-dragover');
    const file = event.dataTransfer?.files?.[0];
    if (file) setFile(file);
  });
  checkButton.addEventListener('click', executeCheck);
  clearButton.addEventListener('click', clearFile);
}

function setFile(file) {
  hideError();
  if (!file) return clearFile();
  if (!/\.(docx|xlsx)$/i.test(file.name)) return showError('対応形式は .docx と .xlsx です。');
  state.file = file;
  state.model = null;
  state.findings = [];
  selectedFile.hidden = false;
  selectedFile.textContent = `${file.name} — ${formatBytes(file.size)}`;
  checkButton.disabled = false;
  clearButton.disabled = false;
  updateResultsUi();
}

function clearFile() {
  state.file = null;
  state.model = null;
  state.findings = [];
  fileInput.value = '';
  selectedFile.hidden = true;
  selectedFile.textContent = '';
  checkButton.disabled = true;
  clearButton.disabled = true;
  progress.hidden = true;
  hideError();
  updateResultsUi();
}

async function executeCheck() {
  if (!state.file) return;
  hideError();
  progress.hidden = false;
  checkButton.disabled = true;
  setProgress(5, '開始しています');
  try {
    const model = await parseOfficeFile(state.file, setProgress);
    state.model = model;
    state.findings = runDocumentQc(model, state.enabledRuleIds, setProgress);
    setProgress(100, `完了: ${state.findings.length}件の指摘`);
    updateResultsUi();
    setTimeout(() => activateView('results'), 250);
  } catch (error) {
    console.error(error);
    showError(error instanceof Error ? error.message : String(error));
  } finally {
    checkButton.disabled = false;
  }
}

function setProgress(percent, text) {
  progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  progressText.textContent = text;
}

function showError(message) {
  errorBox.hidden = false;
  errorBox.textContent = message;
}

function hideError() {
  errorBox.hidden = true;
  errorBox.textContent = '';
}

function bindResults() {
  $('resultSearch').addEventListener('input', renderFindings);
  $('severityFilter').addEventListener('change', renderFindings);
  $('exportCsv').addEventListener('click', () => downloadCsv(state.findings, state.file?.name));
  $('exportXlsx').addEventListener('click', () => downloadXlsx(state.findings, state.file?.name));
  $('resetRules').addEventListener('click', () => {
    state.enabledRuleIds = new Set(RULES.map(rule => rule.id));
    renderRules();
    renderRuleSummary();
  });
}

function updateResultsUi() {
  $('resultCountBadge').textContent = String(state.findings.length);
  const hasRun = Boolean(state.model);
  $('resultMeta').textContent = hasRun
    ? `${state.file?.name || ''} / ${state.model.kind === 'word' ? 'Word' : 'Excel'} / ${state.findings.length}件`
    : 'まだチェックしていません。';
  $('exportCsv').disabled = !hasRun;
  $('exportXlsx').disabled = !hasRun;

  if (!hasRun) {
    $('emptyResults').hidden = true;
    $('resultsWrap').hidden = true;
  } else if (!state.findings.length) {
    $('emptyResults').hidden = false;
    $('resultsWrap').hidden = true;
    $('emptyResults').querySelector('strong').textContent = '指摘はありません';
    $('emptyResults').querySelector('span').textContent = '現在有効なルールでは問題を検出しませんでした。';
  } else {
    $('emptyResults').hidden = true;
    $('resultsWrap').hidden = false;
  }
  renderFindings();
}

function renderFindings() {
  const body = $('resultsBody');
  body.textContent = '';
  const q = $('resultSearch').value.trim().toLowerCase();
  const severity = $('severityFilter').value;
  const rows = state.findings.filter(finding =>
    (severity === 'all' || finding.severity === severity) &&
    (!q || `${finding.ruleId} ${finding.location} ${finding.matched} ${finding.message} ${finding.suggestion} ${finding.context}`.toLowerCase().includes(q))
  );

  for (const finding of rows) {
    const tr = document.createElement('tr');
    tr.append(cell(codeText(finding.ruleId)));
    tr.append(cell(severityBadge(finding.severity)));
    tr.append(cell(finding.location));
    tr.append(cell(finding.matched));
    const messageCell = cell(finding.message);
    messageCell.classList.add('message');
    if (finding.context) {
      const context = document.createElement('div');
      context.className = 'muted';
      context.textContent = finding.context;
      messageCell.append(context);
    }
    tr.append(messageCell);
    tr.append(cell(finding.suggestion));
    body.append(tr);
  }

  if (state.model && state.findings.length && !rows.length) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 6;
    td.textContent = '条件に一致する指摘はありません。';
    td.className = 'muted';
    tr.append(td);
    body.append(tr);
  }
}

function renderRules() {
  const container = $('rulesList');
  container.textContent = '';
  for (const [category, label] of Object.entries(CATEGORY_LABELS)) {
    const rules = RULES.filter(rule => rule.category === category);
    if (!rules.length) continue;
    const group = document.createElement('section');
    group.className = 'rule-group';
    const header = document.createElement('div');
    header.className = 'rule-group-header';
    const h3 = document.createElement('h3');
    h3.textContent = `${label} (${rules.length})`;
    const toggle = document.createElement('button');
    toggle.className = 'button';
    toggle.textContent = '全て切替';
    toggle.addEventListener('click', () => {
      const allOn = rules.every(rule => state.enabledRuleIds.has(rule.id));
      for (const rule of rules) allOn ? state.enabledRuleIds.delete(rule.id) : state.enabledRuleIds.add(rule.id);
      renderRules();
      renderRuleSummary();
    });
    header.append(h3, toggle);
    group.append(header);

    for (const rule of rules) {
      const row = document.createElement('div');
      row.className = 'rule-row';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = state.enabledRuleIds.has(rule.id);
      checkbox.addEventListener('change', () => {
        checkbox.checked ? state.enabledRuleIds.add(rule.id) : state.enabledRuleIds.delete(rule.id);
        renderRuleSummary();
      });
      const id = document.createElement('span');
      id.className = 'rule-id';
      id.textContent = rule.id;
      const desc = document.createElement('div');
      const strong = document.createElement('strong');
      strong.textContent = rule.title;
      const p = document.createElement('p');
      p.textContent = rule.description;
      desc.append(strong, p);
      row.append(checkbox, id, desc, severityBadge(rule.severity));
      group.append(row);
    }
    container.append(group);
  }
}

function renderRuleSummary() {
  const root = $('ruleSummary');
  root.textContent = '';
  const data = [
    ['有効', state.enabledRuleIds.size],
    ['Word', RULES.filter(rule => state.enabledRuleIds.has(rule.id) && rule.appliesTo.includes('word')).length],
    ['Excel', RULES.filter(rule => state.enabledRuleIds.has(rule.id) && rule.appliesTo.includes('excel')).length]
  ];
  for (const [label, value] of data) {
    const div = document.createElement('div');
    div.className = 'stat';
    const strong = document.createElement('strong');
    strong.textContent = value;
    const span = document.createElement('span');
    span.textContent = label;
    div.append(strong, span);
    root.append(div);
  }
}

function severityBadge(severity) {
  const span = document.createElement('span');
  span.className = `severity ${severity}`;
  span.textContent = severity === 'error' ? '要修正' : severity === 'confirm' ? '確認' : '参考';
  return span;
}

function codeText(text) {
  const code = document.createElement('code');
  code.textContent = text;
  return code;
}

function cell(content) {
  const td = document.createElement('td');
  if (content instanceof Node) td.append(content);
  else td.textContent = content ?? '';
  return td;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
}
