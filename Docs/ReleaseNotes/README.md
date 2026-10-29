# リリース本文に足す文章

`v<バージョン>.md`（例: `v2.8.md`）を置くと、リリース本文の「Revit 2021 / … 対応」の行の下に自動で差し込まれる。
自動生成の本文（`scripts/generate-release-body.py`）は新機能（`features.json` の `added_in`）しか載せないため、
**改善・不具合の修正はここに書く**。

- 差し込むのは `scripts/insert-release-notes.py`。`<!-- release-notes:start -->` 〜 `<!-- release-notes:end -->` の間に入れる
- 下書きを作るとき（`build-and-release.yml`）と、公開するとき（`publish-release.yml`）の両方で差し込む。
  公開のときは**このファイルの最新の内容で置き換える**ので、下書きを作った後に直しても、公開前に main に入れれば反映される
- 運営者が GitHub の画面で本文を編集する必要はない
