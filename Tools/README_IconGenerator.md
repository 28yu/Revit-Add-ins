# 🎨 28Tools アイコンジェネレーター

## 概要
Revitスタイルのアイコンを効率的に作成するためのツールです。

## 使用方法

### 1. 基本的な使い方

```bash
# IconGenerator.htmlをブラウザで開く
# Windows: エクスプローラーでダブルクリック
# macOS/Linux: 右クリック → プログラムから開く → ブラウザ
```

### 2. アイコンのダウンロード

1. ブラウザでIconGenerator.htmlを開く
2. プレビューでデザインを確認
3. 「32×32をダウンロード」ボタンをクリック
4. `filled_region_32.png` がダウンロードされます

### 3. カスタムアイコンの作成

IconGenerator.htmlの `createIcon()` 関数を編集して、独自のデザインを作成できます。

## アイコンデザインガイドライン

### 基本仕様
- **サイズ**: 32×32px（標準）、64×64px（高解像度）
- **背景**: 透明
- **フォーマット**: PNG

### カラーパレット
```
Revitブルー（メイン）: #0066CC
アクセント（明るい）:   #3399FF
線（濃いグレー）:      #333333
背景要素（薄いグレー）: #E6E6E6, #F0F0F0
```

### デザインルール
1. **シンプル**: 32×32pxで判別可能な要素サイズ
2. **一貫性**: Revitブルー系で統一
3. **視認性**: 線の太さ 0.8～1.2px
4. **機能表現**: 機能の本質を視覚化

## SVGテンプレート

### 基本構造
```html
<svg width="32" height="32" viewBox="0 0 32 32">
    <!-- 背景（透明） -->

    <!-- メイン要素 -->
    <rect x="6" y="6" width="20" height="20"
          fill="#3399FF"
          stroke="#333333"
          stroke-width="1.2"/>

    <!-- アクセント要素 -->
    <path d="..."
          stroke="#0066CC"
          stroke-width="1.5"
          fill="none"/>
</svg>
```

### よく使う要素

#### 矩形
```javascript
const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
rect.setAttribute('x', 10);
rect.setAttribute('y', 10);
rect.setAttribute('width', 12);
rect.setAttribute('height', 12);
rect.setAttribute('fill', '#3399FF');
rect.setAttribute('stroke', '#333333');
rect.setAttribute('stroke-width', '1.2');
```

#### 円
```javascript
const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
circle.setAttribute('cx', 16);
circle.setAttribute('cy', 16);
circle.setAttribute('r', 8);
circle.setAttribute('fill', '#0066CC');
circle.setAttribute('stroke', '#333333');
circle.setAttribute('stroke-width', '1.2');
```

#### パス（矢印など）
```javascript
const arrow = document.createElementNS('http://www.w3.org/2000/svg', 'path');
arrow.setAttribute('d', 'M 10 16 L 22 16 M 22 16 L 18 12 M 22 16 L 18 20');
arrow.setAttribute('stroke', '#333333');
arrow.setAttribute('stroke-width', '1.5');
arrow.setAttribute('fill', 'none');
arrow.setAttribute('stroke-linecap', 'round');
```

## アイコンの配置

### リポジトリへの追加
```bash
# ダウンロードしたPNGを配置
cp filled_region_32.png ../Resources/Icons/

# Tools28.csprojに追加（既に追加済みの場合は不要）
# <Resource Include="Resources\Icons\filled_region_32.png" />
```

### Application.csでの使用
```csharp
filledRegionButtonData.LargeImage = LoadImage("filled_region_32.png");
```

## トラブルシューティング

### アイコンがぼやける
- 32×32pxで作成しているか確認
- 線の太さが細すぎないか確認（最低0.8px）

### ダウンロードできない
- ブラウザのポップアップブロッカーを無効化
- 別のブラウザで試す（Chrome推奨）

### デザインが反映されない
- ブラウザのキャッシュをクリア
- HTMLファイルを再読み込み（Ctrl+F5）

## 効率化のヒント

### 1. テンプレートの複製
新しいアイコンを作成する際は、IconGenerator.htmlをコピーして、`createIcon()` 関数のみ編集します。

### 2. バッチ生成
複数のアイコンを一度に生成したい場合は、HTMLに複数のSVG要素を追加します。

### 3. デザインツールとの連携
- Figmaで下書き → SVGパスをコピー → HTMLに貼り付け
- Illustratorで作成 → SVGでエクスポート → パスを抽出

## 参考リンク
- [SVG Path Reference](https://developer.mozilla.org/ja/docs/Web/SVG/Tutorial/Paths)
- [Revit API Icon Guidelines](https://www.autodesk.com/developer-network/platform-technologies/revit)
