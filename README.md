# document-qc

Word (`.docx`) / Excel (`.xlsx`) を**ブラウザ内だけで**解析する、完全静的な品質チェックアプリです。

## 方針

- サーバーAPIなし
- アップロード処理なし（File APIで端末内読込）
- 外部CDNなし
- Analytics / Sentry / 広告タグなし
- `Content-Security-Policy: connect-src 'none'`
- `fetch` / XHR / WebSocket / EventSource / `sendBeacon` を実行時ガードで無効化
- 文書内容を localStorage / IndexedDB に保存しない
- 組織固有の用語・基準・案件情報を含めない

GitHub Pages等を利用する場合も、ホスティング側はHTML/CSS/JavaScriptを配信するだけです。選択したWord/Excelのバイト列を送信するコードはありません。

> この説明は本アプリケーションの実装境界について述べるものです。ブラウザ、OS、拡張機能、端末上の他ソフトウェアまで含む完全な機密性を保証するものではありません。

## ルール設計

現在は60件の汎用ルールを実装しています。明確な機械判定ができるものは「要修正」、表記方針・簡潔さ・公用文慣行など文脈依存性が残るものは原則「確認」としています。

組織固有の表記基準、案件固有の正式名称、意味論的に区別が必要な語（例：技術用語として正しい「消化」など）は、汎用ルールとして自動置換しません。

## 主な機能

### Word / 文字・表記

- NBSP、連続空白、句読点重複
- 句読点の混在
- 常体 / 敬体の簡易混在判定
- 明確な誤記候補（例：発砲スチロール）
- SDGs / MaaS / NOx / kW 等の大文字小文字
- 公用文・ビジネス文書で確認しやすい表記候補（の通り、予め、出来る等）
- 「約…くらい」「各…ごとに」等の重複表現
- 二重否定、冗長表現の確認
- `% / ％`、`CO2 / CO₂`、スラッシュ、和暦 / 西暦等の混在
- 全角英字で記載された単位

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
- hidden / veryHidden シート
- 非表示行・列
- externalLinks
- 数式領域への直値混入（近傍ヒューリスティック）
- 数式の相対参照パターン異常（近傍ヒューリスティック）
- 単純な自己参照
- 存在しないシート名への数式参照
- Wordと共通の表記ルールの一部をセル文字列にも適用

### 共通

- ルールON/OFF
- 重要度（要修正 / 確認 / 参考）
- 結果検索・絞り込み
- CSV / XLSX エクスポート

## アーキテクチャ

```text
.docx / .xlsx
      │ File API
      ▼
 ArrayBuffer
      ▼
 ZIP parser (browser native DecompressionStream)
      ▼
 OOXML parser (DOMParser)
      ▼
 common document model
      ▼
 deterministic rule engine
      ▼
 results → CSV / XLSX
```

外部ライブラリを実行時にロードしません。ZIPコンテナの読取と、結果XLSXの生成もリポジトリ内のコードだけで行います。

## 対応ブラウザ

`DecompressionStream('deflate-raw')` が利用できる比較的新しいブラウザを対象とします。Chrome / Edge系を主対象としています。

## ローカル実行

ES Modulesを利用するため、`file://` 直開きではなく静的HTTPサーバーから開いてください。

```bash
python -m http.server 8000
```

その後 `http://localhost:8000/` を開きます。

## GitHub Pages

静的ファイルのみなのでGitHub Pagesで公開できます。`.github/workflows/pages.yml` は `main` 更新時にPagesへデプロイする構成です。

## AIについて

現時点では再現性と監査性を優先し、意味論的AIレビューを実装していません。将来Gemini Nano等のオンデバイスAIを追加する場合も、クラウドAPIへのフォールバックは行わず、ファイル内容を外部送信しない設計を維持します。

## ライセンス

MIT
