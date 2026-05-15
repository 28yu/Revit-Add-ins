# 28tools-download サイト 実装指示書

> このドキュメントは `28yu/28tools-download` リポジトリで実装すべき内容を記載しています。本ファイル自体は Revit-Add-ins リポジトリで管理されています。

## 目的

配布サイト（https://28yu.github.io/28tools-download/）の機能紹介ページとマニュアルページを、Revit-Add-ins リポジトリの `Docs/features.json` を**唯一の正本**として動的生成する。

## アーキテクチャ

```
[Revit-Add-ins リポジトリ]                          [28tools-download リポジトリ]
                                                
Docs/features.json ──────────┐                    index.html (機能カタログ)
Docs/Features/*.md ──────────┼──── fetch ────►   manual.html (マニュアル表示)
Docs/icons/features/*.png ───┘                    
                                                  ※ JS で fetch して動的描画
                                                  ※ 配布サイト側はコード変更不要
                                                    （新機能追加時も）
```

## エンドポイント（Revit-Add-ins 側）

main ブランチ反映後に以下が GitHub Pages から取得可能：

| 種類 | URL |
|------|-----|
| 機能カタログ | `https://28yu.github.io/Revit-Add-ins/features.json` |
| アイコン | `https://28yu.github.io/Revit-Add-ins/icons/features/<file>.png` |
| マニュアル | `https://28yu.github.io/Revit-Add-ins/Features/<id>.md` |

## features.json のスキーマ

```json
{
  "version": 1,
  "baseUrl": "https://28yu.github.io/Revit-Add-ins",
  "categories": {
    "<cat_id>": { "ja": "...", "en": "...", "zh": "..." }
  },
  "features": [
    {
      "id": "FeatureId",
      "category": "<cat_id>",
      "icon": "icons/features/xxx.png",        // baseUrl からの相対パス
      "manual": "Features/Xxx.md",             // baseUrl からの相対パス
      "added_in": "2.1",                        // 任意。新機能バッジ判定用
      "names": { "ja": "...", "en": "...", "zh": "..." }
    }
  ]
}
```

---

## 実装指示

### 1. 機能カタログページ（index.html or features.html）

機能カードを動的生成する。以下のサンプルコードをサイトに組み込む。

```html
<!-- 言語切替 UI（既存があれば流用） -->
<div class="lang-switcher">
  <button data-lang="ja">日本語</button>
  <button data-lang="en">English</button>
  <button data-lang="zh">中文</button>
</div>

<!-- カテゴリごとの機能カードがここに描画される -->
<div id="feature-catalog"></div>

<script>
const CATALOG_URL = 'https://28yu.github.io/Revit-Add-ins/features.json';
let catalogData = null;
let currentLang = localStorage.getItem('lang') || 'ja';

async function loadCatalog() {
  const res = await fetch(CATALOG_URL);
  catalogData = await res.json();
  render();
}

function render() {
  const { baseUrl, categories, features } = catalogData;
  const container = document.getElementById('feature-catalog');
  container.innerHTML = '';

  // カテゴリ順は features.json の categories キー順
  for (const catId of Object.keys(categories)) {
    const catFeatures = features.filter(f => f.category === catId);
    if (catFeatures.length === 0) continue;

    const section = document.createElement('section');
    section.className = 'feature-category';
    section.innerHTML = `<h2>${categories[catId][currentLang]}</h2>`;

    const grid = document.createElement('div');
    grid.className = 'feature-grid';

    for (const f of catFeatures) {
      const card = document.createElement('a');
      card.className = 'feature-card';
      card.href = `manual.html?id=${f.id}`;
      card.innerHTML = `
        <img src="${baseUrl}/${f.icon}" alt="${f.names[currentLang]}" />
        <div class="feature-name">${f.names[currentLang]}</div>
      `;
      grid.appendChild(card);
    }
    section.appendChild(grid);
    container.appendChild(section);
  }
}

document.querySelectorAll('.lang-switcher button').forEach(btn => {
  btn.addEventListener('click', () => {
    currentLang = btn.dataset.lang;
    localStorage.setItem('lang', currentLang);
    render();
  });
});

loadCatalog();
</script>

<style>
.feature-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
  gap: 16px;
}
.feature-card {
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 16px;
  border: 1px solid #ddd;
  border-radius: 8px;
  text-decoration: none;
  color: inherit;
  transition: box-shadow .2s;
}
.feature-card:hover {
  box-shadow: 0 4px 12px rgba(0,0,0,.1);
}
.feature-card img {
  width: 64px;
  height: 64px;
  object-fit: contain;
}
.feature-name {
  margin-top: 8px;
  font-size: 14px;
  text-align: center;
}
</style>
```

### 2. マニュアル表示ページ（manual.html）

クエリパラメータ `?id=FeatureId` で機能を指定し、対応する Markdown を取得して HTML として描画する。

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>マニュアル - 28 Tools</title>
  <link rel="stylesheet" href="manual.css">
  <!-- Markdown レンダラ (CDN) -->
  <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
</head>
<body>
  <a href="index.html">← 機能一覧に戻る</a>
  <article id="manual-content">読み込み中...</article>

<script>
const CATALOG_URL = 'https://28yu.github.io/Revit-Add-ins/features.json';
const params = new URLSearchParams(location.search);
const featureId = params.get('id');

async function loadManual() {
  if (!featureId) {
    document.getElementById('manual-content').textContent = '機能IDが指定されていません。';
    return;
  }

  const catalog = await fetch(CATALOG_URL).then(r => r.json());
  const feature = catalog.features.find(f => f.id === featureId);

  if (!feature) {
    document.getElementById('manual-content').textContent = `機能「${featureId}」が見つかりません。`;
    return;
  }

  const mdUrl = `${catalog.baseUrl}/${feature.manual}`;
  const md = await fetch(mdUrl).then(r => r.text());
  document.getElementById('manual-content').innerHTML = marked.parse(md);

  // ページタイトル更新
  document.title = `${feature.names.ja} - 28 Tools`;
}

loadManual();
</script>
</body>
</html>
```

### 3. CSS（マニュアルの見た目）

```css
/* manual.css */
#manual-content {
  max-width: 800px;
  margin: 0 auto;
  padding: 24px;
  line-height: 1.7;
}
#manual-content h1 { border-bottom: 2px solid #333; padding-bottom: 8px; }
#manual-content h2 { margin-top: 32px; border-left: 4px solid #0066cc; padding-left: 8px; }
#manual-content table {
  border-collapse: collapse;
  margin: 16px 0;
}
#manual-content th, #manual-content td {
  border: 1px solid #ccc;
  padding: 6px 12px;
}
#manual-content code {
  background: #f4f4f4;
  padding: 2px 6px;
  border-radius: 3px;
}
```

---

## 移行ステップ

1. **既存の機能紹介ページのバックアップ** — `index-old.html` などに退避
2. **`index.html` を上記サンプルで置き換え** または 既存ページに `#feature-catalog` 要素と JS を埋め込む
3. **`manual.html` と `manual.css` を新規追加**
4. **ローカル動作確認**: `python3 -m http.server` で起動し、`features.json` が CORS エラーなく取得できることを確認（GitHub Pages 同士は問題なし）
5. **デプロイ**: `28tools-download` の main ブランチに push

## 確認項目

- [ ] `features.json` が取得できる（DevTools の Network タブで 200 OK）
- [ ] 7カテゴリ × 14機能のカードが表示される
- [ ] 言語切替で名称が変わる（JP/EN/CN）
- [ ] カードクリックで対応する Markdown が表示される
- [ ] 画像（アイコン）が表示される

## 新機能追加時の運用

**配布サイト側で行う作業: ゼロ。**

1. Revit-Add-ins 側で `features.json` に新機能エントリを追加
2. `Docs/Features/NewFeature.md` を追加
3. main ブランチに反映 → GitHub Pages 自動デプロイ
4. → 配布サイトを開くと**自動的に新機能カードが追加される**

## トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| カードが表示されない | features.json の取得失敗 | DevTools で URL 確認・CORS 確認 |
| アイコンが表示されない | deploy-pages ワークフロー未実行 | Revit-Add-ins の Actions タブで確認 |
| マニュアルが「読み込み中」のまま | MDファイルパスが間違い | features.json の `manual` フィールド確認 |
| 文字化け | charset 未指定 | HTML に `<meta charset="utf-8">` |

## 参考リンク

- features.json の正本: https://github.com/28yu/Revit-Add-ins/blob/main/Docs/features.json
- マニュアル本文: https://github.com/28yu/Revit-Add-ins/tree/main/Docs/Features
- アイコン: https://github.com/28yu/Revit-Add-ins/tree/main/Resources/Icons
- Pages デプロイ: `.github/workflows/deploy-pages.yml`
