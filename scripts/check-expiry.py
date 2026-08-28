#!/usr/bin/env python3
"""Licensing/ExpiryManager.cs の ExpiryDate を検査する。

リリース時に有効期限の更新を忘れると、配布直後に期限切れになる版が出回るため、
CI で残り日数を確認する。

  --warn-days N  残り N 日以下なら警告（ビルドは通す）
  --fail-days N  残り N 日以下、または期限切れならビルドを失敗させる
"""

import argparse
import datetime
import pathlib
import re
import sys

SOURCE = pathlib.Path("Licensing/ExpiryManager.cs")
PATTERN = re.compile(r"ExpiryDate\s*=\s*new\s+DateTime\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--warn-days", type=int, default=180)
    parser.add_argument("--fail-days", type=int, default=30)
    args = parser.parse_args()

    if not SOURCE.exists():
        print(f"::error::{SOURCE} が見つかりません")
        return 1

    match = PATTERN.search(SOURCE.read_text(encoding="utf-8"))
    if not match:
        print(f"::error::{SOURCE} から ExpiryDate を読み取れませんでした")
        return 1

    expiry = datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    remaining = (expiry - datetime.date.today()).days
    print(f"ExpiryDate = {expiry.isoformat()} (残り {remaining} 日)")

    if remaining <= args.fail_days:
        print(
            f"::error::有効期限が残り {remaining} 日です。"
            f"Licensing/ExpiryManager.cs の ExpiryDate をリリース日の1年後に更新してください"
        )
        return 1

    if remaining <= args.warn_days:
        print(
            f"::warning::有効期限が残り {remaining} 日です。"
            f"このリリースに合わせて ExpiryDate の更新を検討してください"
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
