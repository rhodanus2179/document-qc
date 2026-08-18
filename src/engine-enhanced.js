import { runRules as runBaseRules } from './engine.js';
import { RULES, EXTRA_RULES } from './rules-enhanced.js';
import { PUBLIC_DOCUMENTS } from './public-master.js';

export function runRules(model, enabledIds = new Set(RULES.map(r => r.id)), onProgress = () => {}) {
  const baseFindings = runBaseRules(model, enabledIds, onProgress)
    .filter(f => f.ruleId !== 'TXT-005');

  const findings = [...baseFindings];

  if (model.kind === 'word' && enabledIds.has('TXT-005')) {
    const rule = RULES.find(r => r.id === 'TXT-005');
    const finding = sentenceRegisterMix(model, rule);
    if (finding) findings.push(finding);
  }

  const applicable = EXTRA_RULES.filter(r => enabledIds.has(r.id) && r.appliesTo.includes(model.kind));
  applicable.forEach((rule, index) => {
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
    runExtra(rule, model, push);
    onProgress(96 + Math.round(((index + 1) / Math.max(1, applicable.length)) * 3), `追加ルール ${index + 1}/${applicable.length}: ${rule.id}`);
  });

  return findings;
}

function runExtra(rule, model, push) {
  switch (rule.engine) {
    case 'source-typo-context': return sourceTypoContext(model, push);
    case 'pruning-branch-context': return pruningBranchContext(model, push);
    case 'discrepancy-typo-context': return discrepancyTypoContext(model, push);
    case 'literal-extra': return literalExtra(rule, model, push);
    case 'public-official-name': return publicOfficialName(model, push);
    case 'public-decision-date': return publicDecisionDate(model, push);
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

function sentenceRegisterMix(model, rule) {
  let polite = 0;
  let plain = 0;

  for (const p of model.paragraphs) {
    for (const sentence of splitSentences(p.text)) {
      const clean = sentence.trim().replace(/[。．！？!?]+$/g, '').trim();
      if (!clean) continue;
      if (/(です|ます|でした|ません|ございます)$/.test(clean)) polite++;
      if (/(である|だった|であった|となる|とする)$/.test(clean)) plain++;
    }
  }

  if (!polite || !plain) return null;
  const preferred = polite >= plain ? '敬体（です・ます調）' : '常体（である調）';
  return {
    ruleId: rule.id,
    category: rule.category,
    severity: rule.severity,
    title: rule.title,
    location: '文書全体',
    matched: `敬体 ${polite} / 常体 ${plain}`,
    message: `文体が混在しています。文単位の簡易判定では ${preferred} が多数です。`,
    suggestion: '引用等を除き、文書方針に合わせて確認',
    context: '段落末尾だけでなく、句点等で区切った各文の終止表現を集計しています。'
  };
}

function splitSentences(text) {
  return text.match(/[^。．！？!?]+[。．！？!?]?/g) || [];
}

function sourceTypoContext(model, push) {
  const re = /(資料|図表|図|表|データ|統計|文献|参考資料|引用元)(?:\s*の)?\s*出展/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      const offset = m[0].lastIndexOf('出展');
      const start = m.index + offset;
      push({
        location: item.location,
        matched: '出展',
        message: '資料・図表等の出所を示す文脈で「出展」が使われています。「出典」の誤記でないか確認してください。',
        suggestion: '出典',
        context: item.context(start, 2)
      });
    }
  }
}

function pruningBranchContext(model, push) {
  const domain = /(処理|収集|廃棄|資源化|焼却|堆肥|造園|剪定|枝葉|チップ|破砕)/;
  for (const item of contentItems(model)) {
    let start = 0;
    while ((start = item.text.indexOf('選定枝', start)) !== -1) {
      const windowStart = Math.max(0, start - 24);
      const windowEnd = Math.min(item.text.length, start + 3 + 24);
      const windowText = item.text.slice(windowStart, windowEnd);
      if (domain.test(windowText)) {
        push({
          location: item.location,
          matched: '選定枝',
          message: '枝葉の処理・造園等の文脈です。「剪定枝」の誤記でないか確認してください。',
          suggestion: '剪定枝',
          context: item.context(start, 3)
        });
      }
      start += 3;
    }
  }
}

function discrepancyTypoContext(model, push) {
  const re = /祖語(?=[^。．！？!?\n]{0,24}(?:生じ|生ず|解消|防止|避け|ないよう|なく))/g;
  for (const item of contentItems(model)) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(item.text))) {
      push({
        location: item.location,
        matched: '祖語',
        message: '食い違い・不一致が生じる文脈です。「齟齬」の誤記でないか確認してください。',
        suggestion: '齟齬',
        context: item.context(m.index, 2)
      });
    }
  }
}

function literalExtra(rule, model, push) {
  for (const replacement of rule.replacements || []) {
    for (const item of contentItems(model)) {
      let start = 0;
      while ((start = item.text.indexOf(replacement.from, start)) !== -1) {
        push({
          location: item.location,
          matched: replacement.from,
          message: `「${replacement.from}」の表記を確認してください。`,
          suggestion: replacement.to,
          context: item.context(start, replacement.from.length)
        });
        start += Math.max(1, replacement.from.length);
      }
    }
  }
}

function publicOfficialName(model, push) {
  for (const master of PUBLIC_DOCUMENTS) {
    for (const item of contentItems(model)) {
      master.namePattern.lastIndex = 0;
      let m;
      while ((m = master.namePattern.exec(item.text))) {
        if (m[0] === master.officialName) continue;
        push({
          location: item.location,
          matched: m[0],
          message: `公開情報上の正式名称は「${master.officialName}」です。`,
          suggestion: master.officialName,
          context: item.context(m.index, m[0].length)
        });
      }
    }
  }
}

function publicDecisionDate(model, push) {
  for (const master of PUBLIC_DOCUMENTS) {
    for (const item of contentItems(model)) {
      master.namePattern.lastIndex = 0;
      let planMatch;
      while ((planMatch = master.namePattern.exec(item.text))) {
        const a = Math.max(0, planMatch.index - 16);
        const b = Math.min(item.text.length, planMatch.index + planMatch[0].length + 48);
        const windowText = item.text.slice(a, b);
        const era = /令和[\s　]*([0-9０-９]+)[\s　]*年/.exec(windowText);
        const western = /((?:19|20)[0-9０-９]{2})[\s　]*年/.exec(windowText);
        let wrong = null;
        if (era && Number(toAsciiDigits(era[1])) !== master.cabinetDecision.japaneseEraYear) wrong = era[0];
        if (!wrong && western && Number(toAsciiDigits(western[1])) !== master.cabinetDecision.westernYear) wrong = western[0];
        if (!wrong) continue;
        const wrongIndex = item.text.indexOf(wrong, a);
        push({
          location: item.location,
          matched: wrong,
          message: `「${master.officialName}」は${master.cabinetDecision.display}に閣議決定されています。近接する年表記と一致しません。`,
          suggestion: master.cabinetDecision.display,
          context: item.context(wrongIndex >= 0 ? wrongIndex : planMatch.index, wrong.length)
        });
      }
    }
  }
}

function paragraphLocation(p) {
  return `${p.pageHint ? `保存時ページ目安 ${p.pageHint} / ` : ''}第${p.index}段落`;
}

function excerpt(text, start, length) {
  const a = Math.max(0, start - 24);
  const b = Math.min(text.length, start + length + 36);
  return `${a ? '…' : ''}${text.slice(a, b)}${b < text.length ? '…' : ''}`;
}

function toAsciiDigits(s) {
  return s.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
}
