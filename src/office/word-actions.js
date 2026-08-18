const SEARCH_OPTIONS = {
  matchCase: true,
  matchWholeWord: false,
  matchWildcards: false,
  matchPrefix: false,
  matchSuffix: false,
  ignorePunct: false,
  ignoreSpace: false
};

export async function locateFinding(finding) {
  if (!finding?.anchor || finding.anchor.kind !== 'word-text') throw new Error('この指摘には移動可能な位置情報がありません。');
  await Word.run(async context => {
    const target = await resolveFindingRange(context, finding);
    target.select();
    await context.sync();
  });
}

export async function applyFindingFix(finding) {
  if (!finding?.fix || !finding?.anchor) throw new Error('この指摘は自動修正の対象ではありません。');
  await applyFindingFixes([finding]);
  return 1;
}

export async function applyFindingFixes(findings) {
  const actionable = dedupeFindings((findings || []).filter(finding => finding?.fix && finding?.anchor?.kind === 'word-text'));
  if (!actionable.length) return 0;

  await Word.run(async context => {
    const body = context.document.body;
    const groups = new Map();

    for (const finding of actionable) {
      const searchText = finding.anchor.searchText;
      if (!groups.has(searchText)) {
        const collection = body.search(toWordSearchText(searchText), SEARCH_OPTIONS);
        collection.load('items/text');
        groups.set(searchText, { collection, findings: [] });
      }
      groups.get(searchText).findings.push(finding);
    }

    await context.sync();

    for (const [searchText, group] of groups) {
      const exact = exactRanges(group.collection.items, searchText);
      for (const finding of group.findings) {
        const range = exact[finding.anchor.occurrence];
        if (!range) throw staleFindingError(finding);
        range.insertText(finding.fix.replacement, Word.InsertLocation.replace);
      }
    }

    await context.sync();
  });

  return actionable.length;
}

async function resolveFindingRange(context, finding) {
  const searchText = finding.anchor.searchText;
  const collection = context.document.body.search(toWordSearchText(searchText), SEARCH_OPTIONS);
  collection.load('items/text');
  await context.sync();
  const exact = exactRanges(collection.items, searchText);
  const range = exact[finding.anchor.occurrence];
  if (!range) throw staleFindingError(finding);
  return range;
}

function exactRanges(items, searchText) {
  if (searchText === '\u00a0') return items.filter(item => item.text === '\u00a0');
  return items.filter(item => item.text === searchText);
}

function toWordSearchText(searchText) {
  // Word's search grammar uses ^s for a nonbreaking space.
  return searchText === '\u00a0' ? '^s' : searchText;
}

function dedupeFindings(findings) {
  const seen = new Set();
  return findings.filter(finding => {
    const key = `${finding.anchor.searchText}\u0000${finding.anchor.occurrence}\u0000${finding.fix.replacement}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function staleFindingError(finding) {
  const error = new Error(`「${finding.title || finding.ruleId}」の該当箇所を現在の文書で再特定できませんでした。文書が変更された可能性があるため、再チェックしてください。`);
  error.code = 'STALE_FINDING';
  return error;
}
