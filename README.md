# NFT Trait Viewer for Google スプレッドシート

Google Apps Script で動作する「NFT Trait Viewer」は、EVMチェーンとXRPLのNFTを取得し、指定したTraitごとにグループ化してGoogleスプレッドシートに一覧表示するツールです。

EVMネットワークではAlchemy API、XRPLネットワークではBithomp APIを利用し、同じインターフェースでNFTデータを取り出せます。

## 主な機能

- EVM / XRPL の両対応
- Alchemy API または Bithomp API 経由でNFTデータを取得
- TraitごとにNFTをグループ化
- 画像・リンク付きでスプレッドシートに出力
- カスタムメニューから簡単操作

## セットアップ手順

1. **Google スプレッドシートを開く**
2. メニューから「拡張機能」→「Apps Script」を選択し、`index.js` の内容を貼り付けて保存
3. スクリプトプロパティを設定する
   - EVM用: `ALCHEMY_API_KEY` にAlchemyのAPI KEYを設定
   - XRPL用: `BITHOMP_API_KEY` にBithompのAPI KEYを設定
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
3. `B1` に `EVM` または `XRPL` を選択し、`B2` にAPIエンドポイントを入力
   - EVM例: `https://eth-mainnet.g.alchemy.com/v2/`
   - XRPL例: `https://bithomp.com/api/v2/nfts`
4. `B3` にウォレットアドレスを入力し、`B4` にコントラクトアドレスまたはXRPLのIssuer/Taxonを入力
   - XRPLでは `issuer` / `taxon` を `issuer/taxon` の形式で指定可能です
5. `A6` 以降に表示したいTrait名を1つずつ入力
6. 「2. Fetch NFT Data」をクリックすると、NFTデータが新しいシートに出力されます

## 注意事項

- Alchemy APIの利用にはAPIキーが必要です
- XRPLはBithomp APIを利用するため、`BITHOMP_API_KEY` が必要です
- Bithomp APIのFREEプランでは1回の取得件数が 100 件までに制限されます。より多くのNFTを取得したい場合は、有料プランへのアップグレードが必要です
- 取得できるNFTはEVMのERC-721/1155などAlchemyが対応しているもの、またはXRPLのBithomp APIが対応しているものに限ります
- 画像やリンクはOpenSeaやBithomp側の仕様変更等により表示できない場合があります

## AI 利用

- プロトタイプは Gemini CLI によって作成しました
- Github Copilot を用いながら人間の手で大部分を修正しています
- XRPL 対応は Gemini が中心です
- README.md は主に Github Copilot により生成されました

## ライセンス

MIT License
