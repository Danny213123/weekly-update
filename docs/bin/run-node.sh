#!/bin/sh
set -eu

if [ -x "/opt/homebrew/opt/node@20/bin/node" ]; then
  NODE_BIN="/opt/homebrew/opt/node@20/bin/node"
elif [ -x "/usr/local/opt/node@20/bin/node" ]; then
  NODE_BIN="/usr/local/opt/node@20/bin/node"
elif command -v node >/dev/null 2>&1; then
  NODE_BIN="$(command -v node)"
else
  echo "Node.js was not found on PATH." >&2
  exit 1
fi

exec "$NODE_BIN" "$@"
