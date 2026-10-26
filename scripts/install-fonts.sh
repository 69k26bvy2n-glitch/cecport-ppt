#!/usr/bin/env bash
# 安装技能内置 MiSans 字体到当前用户（macOS，无需 sudo）
set -euo pipefail

SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)/assets/fonts"
DEST="$HOME/Library/Fonts"

if [ ! -d "$SRC" ]; then
  echo "未找到字体目录：$SRC" >&2
  exit 1
fi

mkdir -p "$DEST"
cp -f "$SRC"/*.ttf "$DEST"/
echo "已安装到 $DEST：$(ls "$SRC" | tr '\n' ' ')"
echo "若 PowerPoint 仍在运行，请重启后字体现生效。"
