#!/usr/bin/env python3
"""リリース本文に、Docs/ReleaseNotes/v<バージョン>.md の文章を差し込む。

使い方: insert-release-notes.py 本文.md 足す文章.md   （本文.md を書き換える）

- 「… 対応」の行（Revit のバージョンの行）の下に入れる。その行が無ければ先頭に入れる
- 印（<!-- release-notes:start --> 〜 end）で囲むので、何度実行しても重ならず、最新の内容に置き換わる
- 足す文章のファイルが無ければ何もしない
"""
import re
import sys
from pathlib import Path

START, END = "<!-- release-notes:start -->", "<!-- release-notes:end -->"


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    body_path, notes_path = Path(sys.argv[1]), Path(sys.argv[2])
    if not notes_path.exists():
        print(f"{notes_path} が無いので、本文はそのまま")
        return 0
    body = body_path.read_text(encoding="utf-8").replace("\r\n", "\n")
    notes = notes_path.read_text(encoding="utf-8").replace("\r\n", "\n").strip()
    block = f"{START}\n{notes}\n{END}"
    if START in body and END in body:
        body = re.sub(re.escape(START) + r".*?" + re.escape(END), lambda _: block, body, flags=re.S)
    else:
        lines = body.split("\n")
        at = next((i for i, l in enumerate(lines) if l.rstrip().endswith("対応")), None)
        if at is None:
            body = block + "\n\n" + body
        else:
            lines[at + 1:at + 1] = ["", block]
            body = "\n".join(lines)
    body_path.write_text(body, encoding="utf-8")
    print(f"{notes_path} を差し込みました")
    return 0


if __name__ == "__main__":
    sys.exit(main())
