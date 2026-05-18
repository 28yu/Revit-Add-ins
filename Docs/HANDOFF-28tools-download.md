# 引き継ぎ資料: 28tools-download サイトへの動的カタログ実装

> このファイルは `28yu/28tools-download` リポジトリで作業する別 Claude Code セッションに渡すためのハンドオフ資料です。
> このファイル自体は `28yu/Revit-Add-ins` リポジトリで管理されています。

---

## あなたへの依頼

`28yu/28tools-download` リポジトリ（配布サイト `https://28yu.github.io/28tools-download/`）に、**機能カタログとマニュアルを動的生成する仕組み**を実装してください。

現状、配布サイトはハードコードされた機能紹介のままで、`Revit-Add-ins` リポジトリ側で機能を追加・更新しても反映されません。
この仕組みを入れることで、**今後 `Revit-Add-ins` 側で `features.json` または `.md` を更新すれば、配布サイトのコードを一切触らずに自動反映**されるようになります。

---

## 仕様書（必読）

完全な実装指示・サンプルコード・CSS・トラブルシューティングまで揃った仕様書が以下にあります。**まず最初にこれを取得して熟読してください。**

```
https://raw.githubusercontent.com/28yu/Revit-Add-ins/main/Docs/INTEGRATION-28tools-download.md
```

WebFetch ツールで取得可能です。これが**唯一の正本**なので、実装はこの仕様書通りに行ってください。

---

## エンドポイント（実装で使う URL）

仕様書にも記載されていますが、ここでも明示しておきます。
`Revit-Add-ins` リポジトリは GitHub Pages により以下を公開しています:

| 種類 | URL |
|------|-----|
| 機能カタログ | `https://28yu.github.io/Revit-Add-ins/features.json` |
| マニュアル本文 | `https://28yu.github.io/Revit-Add-ins/Features/<id>.md` |
| アイコン | `https://28yu.github.io/Revit-Add-ins/icons/features/<file>.png` |

各 URL がブラウザで開けることを確認してから実装を始めてください。

---

## 実装する成果物

仕様書 `INTEGRATION-28tools-download.md` の「実装指示」セクションに完全なサンプルコードがあります。以下を作成または既存ファイルへ組み込んでください:

1. **`index.html`**（既存があれば修正、なければ新規）
   - 機能カタログ（カード一覧）を動的生成する JS を組み込む
   - 言語切替（JP/EN/CN）に対応
   - カテゴリ順は `features.json` の `categories` キー順を維持

2. **`manual.html`**（新規）
   - クエリパラメータ `?id=FeatureId` で機能を指定
   - 対応する Markdown を fetch して `marked.js` で HTML レンダリング
   - 「← 機能一覧に戻る」リンクを設置

3. **`manual.css`**（新規）
   - マニュアル表示用のスタイル

> サンプルコードはコピペで動く完成品です。配布サイトのデザインに合わせて調整は必要ですが、ロジックを書き換える必要はありません。

---

## 動作確認方法

実装後、以下を確認してください:

1. **ローカル確認**
   ```bash
   python3 -m http.server 8000
   ```
   ブラウザで `http://localhost:8000/` を開く

2. **DevTools (F12) → Network タブ**
   - `features.json` へのリクエストが **200 OK** で返ること
   - 言語切替時に再描画されること
   - カードクリックで `manual.html?id=XXX` に遷移すること
   - マニュアルページで `.md` が fetch されレンダリングされること

3. **本番確認**
   - `main` ブランチへ push → GitHub Pages 自動デプロイ
   - `https://28yu.github.io/28tools-download/` を開いて同様の動作を確認
   - 強制リロード（Ctrl+F5）でキャッシュをクリアして検証

---

## 既存サイトの取り扱い

- 既存の機能紹介ページに大きな改修を加える前に、**バックアップ（`index-old.html` など）を取ってから作業**してください
- ヘッダー・フッター・ナビゲーション等の共通部分は既存のデザインを尊重し、機能カタログ部分だけを動的化する形が安全です

---

## 完了の定義

以下が満たされた状態をゴールとしてください:

- [ ] `https://28yu.github.io/28tools-download/` を開くと `features.json` を fetch するリクエストが Network タブに見える
- [ ] 7 カテゴリ・14 機能のカードが正しく表示される
- [ ] 言語切替で日本語/英語/中国語にラベルが切り替わる
- [ ] カードをクリックすると対応する Markdown マニュアルが表示される
- [ ] アイコン画像が表示される
- [ ] `Revit-Add-ins` 側で `Docs/Features/FormworkCalculator.md` を編集 → `main` push → 数分後に配布サイトを強制リロードすると変更が反映される（テスト確認）

---

## 注意事項

- **`28tools-download` リポジトリ内のコードのみを変更してください。** `Revit-Add-ins` リポジトリには手を入れないでください
- 新機能を追加するための作業ではなく、**既存の機能を動的化する作業**です。`features.json` の中身は触らない
- CORS は GitHub Pages 同士なので問題ありません（同一オリジンポリシーには引っかかりません）
- 配布サイトのデザイン（色・フォント・レイアウト）は既存サイトのトーンに合わせて調整してください

---

## 参考リンク

- 仕様書（必読）: `https://raw.githubusercontent.com/28yu/Revit-Add-ins/main/Docs/INTEGRATION-28tools-download.md`
- 配布サイト（現状）: `https://28yu.github.io/28tools-download/`
- 機能カタログ JSON: `https://28yu.github.io/28tools-download/` ← 取得対象
- マスター リポジトリ: `https://github.com/28yu/Revit-Add-ins`
