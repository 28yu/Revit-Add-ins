#!/usr/bin/env python3
"""features.json からリリース本文 (Markdown) を生成する。

使い方:
  python3 scripts/generate-release-body.py --version 2.2 > release-body.md
"""

import argparse
import json
import sys
from pathlib import Path


def _version_key(value):
    """"2.10" > "2.9" を正しく比較するため、バージョン文字列を数値タプルにする。"""
    parts = []
    for chunk in str(value).split("."):
        try:
            parts.append(int(chunk))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True, help="リリースバージョン (例: 2.2)")
    parser.add_argument(
        "--since",
        default="",
        help=(
            "前回リリースバージョン (例: 2.2)。指定すると、これより新しい added_in を"
            "すべて新機能として扱う。未リリースのままバージョンが進んだ機能を"
            "取りこぼさないために使う。省略時は added_in が --version と完全一致するものだけ。"
        ),
    )
    parser.add_argument(
        "--features",
        default="Docs/features.json",
        help="features.json のパス",
    )
    args = parser.parse_args()

    features_path = Path(args.features)
    if not features_path.exists():
        print(f"ERROR: {features_path} が見つかりません", file=sys.stderr)
        sys.exit(1)

    data = json.loads(features_path.read_text(encoding="utf-8"))
    version = args.version
    categories = data["categories"]
    features = data["features"]

    lines = []
    lines.append(f"## 28 Tools v{version}")
    lines.append("")
    lines.append("Revit 2021 / 2022 / 2023 / 2024 / 2025 / 2026 対応")
    lines.append("")

    # 新機能セクション
    #  --since あり: 前回リリースより新しい added_in を全て（未リリース分の取りこぼし防止）
    #  --since なし: added_in が今回バージョンと完全一致するもの
    if args.since:
        new_features = [
            f for f in features
            if f.get("added_in") and _version_key(f["added_in"]) > _version_key(args.since)
        ]
    else:
        new_features = [f for f in features if f.get("added_in") == version]
    if new_features:
        lines.append(f"### ⭐ v{version} 新機能")
        lines.append("")
        for f in new_features:
            cat_name = categories[f["category"]]["ja"]
            name = f["names"]["ja"]
            lines.append(f"- **{name}**（{cat_name}）")
        lines.append("")

    # 全機能一覧（カテゴリ別）
    lines.append("### 全機能一覧")
    lines.append("")

    # カテゴリ順を features.json の categories キー順に従う
    cat_order = list(categories.keys())
    for cat_id in cat_order:
        cat_features = [f for f in features if f["category"] == cat_id]
        if not cat_features:
            continue
        cat_name = categories[cat_id]["ja"]
        lines.append(f"**{cat_name}**")
        for f in cat_features:
            name = f["names"]["ja"]
            badge = " ⭐新機能" if f in new_features else ""
            lines.append(f"- {name}{badge}")
        lines.append("")

    print("\n".join(lines))


if __name__ == "__main__":
    main()
