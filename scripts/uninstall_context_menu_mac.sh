#!/bin/bash
# ============================================================
#   法律文件脱敏工具 - macOS 右键菜单卸载
# ============================================================

SERVICES_DIR="$HOME/Library/Services"

echo "========================================"
echo "  法律文件脱敏工具 - macOS 右键菜单卸载"
echo "========================================"
echo ""

# 删除 "用脱敏工具打开"
if [ -d "$SERVICES_DIR/用脱敏工具打开.workflow" ]; then
    rm -rf "$SERVICES_DIR/用脱敏工具打开.workflow"
    echo "[OK] 已移除: 用脱敏工具打开"
else
    echo "[跳过] 用脱敏工具打开（未安装）"
fi

# 删除 "一键脱敏"
if [ -d "$SERVICES_DIR/一键脱敏.workflow" ]; then
    rm -rf "$SERVICES_DIR/一键脱敏.workflow"
    echo "[OK] 已移除: 一键脱敏"
else
    echo "[跳过] 一键脱敏（未安装）"
fi

# 删除 "一键还原"
if [ -d "$SERVICES_DIR/一键还原.workflow" ]; then
    rm -rf "$SERVICES_DIR/一键还原.workflow"
    echo "[OK] 已移除: 一键还原"
else
    echo "[跳过] 一键还原（未安装）"
fi

echo ""
echo "========================================"
echo "  卸载完成！"
echo "========================================"
echo ""
