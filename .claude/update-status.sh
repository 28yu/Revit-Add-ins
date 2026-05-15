#!/bin/bash
# Stop hook: STATUS.md の冒頭に最終セッション情報を自動更新
REPO=/home/user/Revit-Add-ins
cd "$REPO" 2>/dev/null || exit 0

CHANGED=$(git diff --name-only 2>/dev/null; git diff --cached --name-only 2>/dev/null; git ls-files --others --exclude-standard 2>/dev/null)
CHANGED=$(echo "$CHANGED" | grep -v '^$' | grep -v 'STATUS\.md' | sort -u)
[ -z "$CHANGED" ] && exit 0

DATE=$(date '+%Y-%m-%d %H:%M')
FILES=$(echo "$CHANGED" | head -5 | tr '\n' ',' | sed 's/,$//')

TMP=$(mktemp)
printf '## 最終セッション: %s\n変更ファイル: %s\n\n' "$DATE" "$FILES" > "$TMP"
grep -Ev '^## 最終セッション:|^変更ファイル:' STATUS.md >> "$TMP" 2>/dev/null
mv "$TMP" STATUS.md
true
