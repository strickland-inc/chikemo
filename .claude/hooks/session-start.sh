#!/bin/bash
set -euo pipefail

# クラウドセッションでだけ実行する。
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

# CLASPRC_JSON が設定済みで ~/.clasprc.json が無いときだけ書き出す。
# トークンの中身は出力しない。
if [ -n "${CLASPRC_JSON:-}" ] && [ ! -f "$HOME/.clasprc.json" ]; then
  umask 077
  printf '%s' "$CLASPRC_JSON" > "$HOME/.clasprc.json"
fi
