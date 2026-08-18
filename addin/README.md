# Document QC for Word

`addin/manifest.xml` は Document QC の Word Task Pane Add-in 用マニフェストです。Add-in 本体はこのリポジトリの GitHub Pages から HTTPS で配信し、開いている Word 文書の内容は外部へ送信しません。

## 対象

- 正式な完全チェック対象: Word for Windows / Word for Mac
- Word on the web: Task Pane 自体は開けますが、Word 文書を `.docx` (compressed OOXML) として取得する Common API が提供されていないため、v1.1.0 では完全チェックを実行しません。

## テスト導入

### Windows

Microsoft が案内する add-in-only manifest のテスト方法では、ネットワーク共有フォルダーを「信頼済みアドイン カタログ」として登録し、そこへ `manifest.xml` を置いて Word の **ホーム > アドイン > 詳細設定 > 共有フォルダー** から追加します。ネットワーク共有方式はテスト用途であり、本番配布向けではありません。

### Mac

`manifest.xml` を次のフォルダーへコピーして Word を再起動します。

```text
/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef
```

## 社内本番配布

組織内配布は Microsoft 365 管理センターの Integrated Apps / Centralized Deployment を使用する想定です。個別テスト用の sideload と、本番配布の仕組みは分けます。

## ホスティング

マニフェストは次の静的ファイルを参照します。

```text
https://rhodanus2179.github.io/document-qc/addin/taskpane.html
```

Office.js の配信元は Microsoft 公式 CDN のみ許可し、Task Pane の CSP は `connect-src 'none'` としています。
