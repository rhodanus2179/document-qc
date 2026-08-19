# document-qc

**Version 1.1.0**

Word (`.docx`) / Excel (`.xlsx`) の品質を、文書内容を外部送信せずにチェックするクライアントサイドアプリです。v1.1.0 では従来のブラウザ版に加え、現在開いている Word 文書を直接チェックできる Office Add-in 版を追加しました。

## 2つの利用形態

### Standalone

ブラウザで `.docx` / `.xlsx` を選択してチェックします。

- File API で端末内読込
- サーバーAPIなし
- 外部CDNなし
- `connect-src 'none'`
- `fetch` / XHR / WebSocket / EventSource / `sendBeacon` を実行時ガードで無効化
- 文書内容を localStorage / IndexedDB に保存しない

### Word Add-in

Word の校閲タブから Task Pane を開き、現在の Word 文書を直接チェックします。

- Windows / Mac 版 Word では現在の文書を compressed OOXML として取得し、Standalone と同じ parser / rule engine で解析
- 指摘から Word 上の該当箇所へ移動
- 安全性の高い指摘は個別修正・同一ルール一括修正
- 文脈依存の候補は「候補を適用」として明示
- 文書内容をアプリ側の外部サーバーへ送信しない
- Office.js は Microsoft 公式 CDN から読み込む
- Task Pane は `connect-src 'none'`

Word on the web では compressed OOXML の取得に対応していないため、v1.1.0 の完全チェックは Windows / Mac 版 Word を正式対象とします。

Add-in のテスト導入・社内配布方法は [`addin/README.md`](addin/README.md) を参照してください。

## ルール設計

現在は **68件**の汎用・公開情報ベースのルールを実装しています。明確な機械判定ができるものは「要修正」、表記方針・簡潔さ・公用文慣行など文脈依存性が残るものは原則「確認」としています。

組織固有の表記基準、案件固有の正式名称、意味論的に区別が必要な語は、汎用ルールとして自動置換しません。公開されている法令・計画等の正式名称や決定日は `public-master.js` に分離しています。

旧来の公文書・ワープロ系の表記慣行も互換性確認用として一部収録しています。「独立した1桁数字を全角にする」ルールは現在の一般的な文書で必須とはせず「確認」として提示します。

## 主なチェック

### Word / 文字・表記

- NBSP、連続空白、句読点重複
- 句読点、`% / ％`、`CO2 / CO₂`、スラッシュ、和暦 / 西暦等の混在
- 常体 / 敬体の文単位での簡易混在判定
- 明確な誤記候裖、周辺文脈を限定した誤記候補
- SDGs / MaaS / NOx / kW 等の大文字小文字
- 公用文・ビジネス文書の表記候補
- 旧来表記の互換確認（独立した1桁の半角数字 → 全角数字）
- 重複表現、二重否定、冗長表現
- 全角英字単位

### Word構造・成果品確認

- 空表
- 直接指定された赤字
- コメント / 変更履歴
- 参照エラー文字列
- 図表参照とキャプションの簡易整合
- 図表番号の欠番 / 重複
- 見出し先頭・末尾、本文段落末尾の余分な空白

### Excel

- `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A` 等
- 数式内の `#REF!`
- hidden / veryHidden シート、非表示行・列
- externalLinks
- 数式領域への直値混入、相対参照パターン異常
- 単純な自己参照、存在しないシート名への参照
- Word と共通の表記ルールの一部

### 公開情報照合

- 第六次環境基本計画の正式名称
- 閣議決定日との近接年表記の照合

## アーキテクチャ

```text
                         ┌─ Standalone: File API
.docx / .xlsx ───────────┤
                         └─ Word Add-in: Office Common API
                                      │
                                      ▼
                                 ArrayBuffer
                                      │
                                      ▼
                         ZIP / OOXML parser
                                      │
                                      ▼
                         common document model
                                      │
                                      ▼
                              core.js facade
                                      │
                         deterministic rule engine
                         ├─ generic rules
                         └─ public information master
                                      │
                                      ▼
                            finding metadata
                         ├─ Word再特定アンカー
                         └─ 修正可否 / 置換候補
                              │             │
                              ▼             ▼
                       Standalone UI   Word Task Pane
```

`src/core.js` を UI から見た正式な公開窓口とし、parser / rule engine / finding metadata を UI 実装から分離しています。既存のルールエンジンは Standalone / Add-in で共用します。

## ディレクトリ

```text
src/
  core.js                 共通公開API
  ooxml.js                OOXML → common model
  engine*.js              ルールエンジン
  rules*.js               ルール定義
  finding-meta.js         アンカー・修正メタデータ
  standalone/app.js       ブラウザ版UI
  office/                 Word Add-in adapter / UI
addin/
  manifest.xml            Word Add-in manifest
  taskpane.html/css       Task Pane
  assets/                 リボンアイコン
tests/
```

## Standalone のローカル実行

ES Modules を利用するため、`file://` 直開きではなく静的 HTTP サーバーから開いてください。

```bash
python -m http.server 8000
```

その後 `http://localhost:8000/` を開きます。

## GitHub Pages

静的ファイルのみなので GitHub Pages で配信できます。`.github/workflows/pages.yml` は `main` 更新時に Pages へデプロイします。Add-in の Task Pane も同じ Pages 配下から配信します。

## AIについて

v1.1.0 では再現性と監査性を優先し、意味論的AIレビューを実装していません。将来オンデバイスAIを追加する場合も、クラウドAPIへのフォールバックを行わず、文書内容を外部送信しない設計を基本とします。

