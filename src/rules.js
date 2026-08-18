export const CATEGORY_LABELS = {
  text: '文字・表記',
  word: 'Word構造',
  excel: 'Excel',
  safety: '成果品確認'
};

export const RULES = [
  { id: 'TXT-001', category: 'text', severity: 'error', engine: 'nbsp', title: 'NBSP（不可視空白）', description: 'U+00A0 の不可視空白を検出します。', appliesTo: ['word'] },
  { id: 'TXT-002', category: 'text', severity: 'confirm', engine: 'repeated-space', title: '連続する半角空白', description: '文章中の連続する半角空白を確認します。', appliesTo: ['word'] },
  { id: 'TXT-003', category: 'text', severity: 'confirm', engine: 'repeated-punctuation', title: '句読点の重複', description: '「、、」「。。」などの連続を検出します。', appliesTo: ['word'] },
  { id: 'TXT-004', category: 'text', severity: 'confirm', engine: 'punctuation-mix', title: '句読点表記の混在', description: '「、/，」「。/．」が文書内で混在していないか確認します。', appliesTo: ['word'] },
  { id: 'TXT-005', category: 'text', severity: 'confirm', engine: 'register-mix', title: '常体・敬体の混在', description: '「である調」と「です・ます調」の混在を簡易判定します。', appliesTo: ['word'] },

  { id: 'DOC-001', category: 'word', severity: 'confirm', engine: 'empty-table', title: '空の表', description: '内容が空の表が残っていないか確認します。', appliesTo: ['word'] },
  { id: 'DOC-002', category: 'word', severity: 'confirm', engine: 'red-text', title: '赤字の残存', description: '直接指定された赤色文字を検出します。', appliesTo: ['word'] },
  { id: 'DOC-003', category: 'safety', severity: 'confirm', engine: 'comments', title: 'コメントの残存', description: 'Wordコメントが残っていないか確認します。', appliesTo: ['word'] },
  { id: 'DOC-004', category: 'safety', severity: 'confirm', engine: 'tracked-changes', title: '変更履歴の残存', description: '変更履歴（挿入・削除等）が残っていないか確認します。', appliesTo: ['word'] },
  { id: 'DOC-005', category: 'word', severity: 'error', engine: 'broken-reference-text', title: '参照エラー文字列', description: 'Wordの「参照元が見つかりません」等のエラー文字列を検出します。', appliesTo: ['word'] },
  { id: 'DOC-006', category: 'word', severity: 'confirm', engine: 'missing-caption', title: '図表参照とキャプションの不一致', description: '本文が参照する図・表番号に対応するキャプションがあるか簡易確認します。', appliesTo: ['word'] },
  { id: 'DOC-007', category: 'word', severity: 'confirm', engine: 'caption-gap', title: '図表番号の欠番', description: '同一系列の図・表番号が不自然に飛んでいないか確認します。', appliesTo: ['word'] },

  { id: 'XLS-001', category: 'excel', severity: 'error', engine: 'error-cell', title: 'Excelエラー値', description: '#REF!、#DIV/0!、#VALUE!、#N/A 等を検出します。', appliesTo: ['excel'] },
  { id: 'XLS-002', category: 'excel', severity: 'error', engine: 'broken-formula-ref', title: '数式内の #REF!', description: '数式に壊れた参照が残っていないか確認します。', appliesTo: ['excel'] },
  { id: 'XLS-003', category: 'safety', severity: 'confirm', engine: 'hidden-sheet', title: '非表示シート', description: 'hidden / veryHidden のシートを列挙します。', appliesTo: ['excel'] },
  { id: 'XLS-004', category: 'safety', severity: 'confirm', engine: 'hidden-row', title: '非表示行', description: '非表示行が残っていないか確認します。', appliesTo: ['excel'] },
  { id: 'XLS-005', category: 'safety', severity: 'confirm', engine: 'hidden-col', title: '非表示列', description: '非表示列が残っていないか確認します。', appliesTo: ['excel'] },
  { id: 'XLS-006', category: 'excel', severity: 'confirm', engine: 'external-link', title: '外部リンク', description: 'ワークブック内の externalLinks パーツを検出します。', appliesTo: ['excel'] },
  { id: 'XLS-007', category: 'excel', severity: 'confirm', engine: 'formula-value-outlier', title: '数式列・行への直値混入', description: '上下または左右が数式のセルで、中央だけ直値のセルを検出します。', appliesTo: ['excel'] },
  { id: 'XLS-008', category: 'excel', severity: 'confirm', engine: 'formula-pattern-outlier', title: '数式パターンの局所不一致', description: '上下または左右の数式パターンが一致するのに中央だけ異なるセルを検出します。', appliesTo: ['excel'] },
  { id: 'XLS-009', category: 'excel', severity: 'confirm', engine: 'self-reference', title: '単純な自己参照', description: '数式が自身のセルを直接参照している可能性を検出します。', appliesTo: ['excel'] }
];
