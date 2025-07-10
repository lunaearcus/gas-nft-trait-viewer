# NFT Trait Viewer for Google スプレッドシート

Google Apps Script で動作する「NFT Trait Viewer」は、指定したウォレットアドレスとコントラクトアドレスのNFTを取得し、指定したTraitごとにグループ化してGoogleスプレッドシートに一覧表示するツールです。

## 主な機能

- Alchemy API経由でNFTデータを取得
- TraitごとにNFTをグループ化
- 画像・OpenSeaリンク付きでスプレッドシートに出力
- カスタムメニューから簡単操作

## セットアップ手順

1. **Google スプレッドシートを開く**
2. メニューから「拡張機能」→「Apps Script」を選択し、`index.js` の内容を貼り付けて保存
3. スクリプトプロパティに `ALCHEMY_API_KEY` を追加し、値にAlchemyのAPIエンドポイントURL（例: `https://eth-mainnet.g.alchemy.com/v2/xxxxxx`）を設定
4. スプレッドシートを再読み込み

## appsscript.json の権限設定について

このスクリプトでは、外部APIへのリクエストやスプレッドシート操作、カスタムUIの利用のため、以下の権限（OAuthスコープ）が必要です。  
`appsscript.json` には次のように記述します（デフォルトで自動生成されますが、手動で設定する場合の参考にしてください）。

```json
// filepath: /workspaces/lunae/working/p1/appsscript.json
{
  // ...existing code...
  "oauthScopes": [
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

- `script.external_request`: 外部API（Alchemy等）へのアクセス
- `spreadsheets`: スプレッドシートの読み書き
- `script.container.ui`: カスタムメニューやダイアログの表示

## 使い方

1. スプレッドシートを開くと「NFT Viewer」メニューが追加されます
2. 「1. Setup Config Sheet」をクリックし、設定用シートを作成
3. `B1` にウォレットアドレス、`B2` にコントラクトアドレスを入力
4. `A5` 以降に表示したいTrait名を1つずつ入力
5. 「2. Fetch NFT Data」をクリックすると、NFTデータが新しいシートに出力されます

## 注意事項

- Alchemy APIの利用にはAPIキーが必要です
- 取得できるNFTはERC-721/1155などAlchemyが対応しているものに限ります
- 画像やリンクはOpenSeaの仕様変更等により表示できない場合があります

## AI 利用

- プロトタイプは Gemini CLI によって作成しました
- Github Copilot を用いながら人間の手で大部分を修正しています
- README.md は主に Github Copilot により生成されました

## ライセンス

MIT License
