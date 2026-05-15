#!/usr/bin/env python3
"""features.json からリリース本文 (Markdown) を生成する。

使い方:
  python3 scripts/generate-release-body.py --version 2.2 > release-body.md
"""

import argparse
import json
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True, help="リリースバージョン (例: 2.2)")
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

    # 新機能セクション（added_in が今回バージョンと一致するもの）
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
            badge = " ⭐新機能" if f.get("added_in") == version else ""
            lines.append(f"- {name}{badge}")
        lines.append("")

    print("\n".join(lines))


if __name__ == "__main__":
    main()
