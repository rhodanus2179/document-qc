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

## 主な機能

### Word

- NBSP、連続空白、句読点重複
- 句読点の混在
- 常体 / 敬体の簡易混在判定
- 空表
- 直接指定された赤字
- コメント / 変更履歴
- 参照エラー文字列
- 図表参照とキャプションの簡易整合
- 図表番号の欠番

### Excel

- `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A` 等
- 数式内の `#REF!`
- hidden / veryHidden シート
- 非表示行・列
- externalLinks
- 数式領域への直値混入（近傍ヒューリスティック）
- 数式の相対参照パターン異常（近傍ヒューリスティック）
- 単純な自己参照

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

初期版は再現性と監査性を優先し、意味論的AIレビューを実装していません。将来Gemini Nano等のオンデバイスAIを追加する場合も、クラウドAPIへのフォールバックは行わず、ファイル内容を外部送信しない設計を維持します。

## ライセンス

MIT
