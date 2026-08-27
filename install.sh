#!/usr/bin/env sh
set -eu
target="${CODEX_HOME:-$HOME/.codex}/skills/map-legal-provision-structure-skill"
mkdir -p "$(dirname "$target")"
if [ -e "$target" ]; then
  echo "Target already exists: $target" >&2
  exit 2
fi
cp -R "$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)" "$target"
echo "Installed: $target"
