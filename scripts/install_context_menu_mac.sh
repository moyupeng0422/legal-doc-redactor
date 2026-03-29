#!/bin/bash
# ============================================================
#   法律文件脱敏工具 - macOS 右键菜单安装
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
SERVICES_DIR="$HOME/Library/Services"

# ---- 颜色定义 ----
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
NC='\033[0m'

echo "========================================"
echo "  法律文件脱敏工具 - macOS 右键菜单安装"
echo "========================================"
echo ""

# ---- 1. 文件检查 ----
if [ ! -f "$SCRIPT_DIR/context_menu_server.py" ]; then
    echo -e "${RED}[错误] 未找到 context_menu_server.py${NC}"
    echo "请确保此脚本位于 scripts/ 目录中"
    exit 1
fi
if [ ! -f "$SCRIPT_DIR/quick_redact.py" ]; then
    echo -e "${RED}[错误] 未找到 quick_redact.py${NC}"
    echo "请确保此脚本位于 scripts/ 目录中"
    exit 1
fi
if [ ! -f "$SCRIPT_DIR/quick_restore.py" ]; then
    echo -e "${RED}[错误] 未找到 quick_restore.py${NC}"
    echo "请确保此脚本位于 scripts/ 目录中"
    exit 1
fi
if [ ! -f "$PROJECT_DIR/assets/index.html" ]; then
    echo -e "${RED}[错误] 未找到 assets/index.html${NC}"
    echo "请确保目录结构完整"
    exit 1
fi
echo -e "${GREEN}[OK]${NC} 文件检查通过"

# ---- 2. 检测 Python3 ----
echo ""
echo -e "${YELLOW}[1/4]${NC} 检测 Python3..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}[错误] 未检测到 python3${NC}"
    echo "请先安装 Python 3: brew install python3"
    echo "或从 https://www.python.org/downloads/ 下载"
    exit 1
fi
PYTHON3_VERSION=$(python3 --version 2>&1)
echo -e "${GREEN}[OK]${NC} $PYTHON3_VERSION"

# ---- 3. 检测/安装 python-docx ----
echo ""
echo -e "${YELLOW}[2/4]${NC} 检测 python-docx..."
if ! python3 -c "import docx" &> /dev/null; then
    echo "python-docx 未安装，正在自动安装..."
    pip3 install python-docx 2>&1 | tail -1
    if ! python3 -c "import docx" &> /dev/null; then
        echo -e "${RED}[错误] python-docx 安装失败，请手动运行: pip3 install python-docx${NC}"
        exit 1
    fi
    echo -e "${GREEN}[OK]${NC} python-docx 安装成功"
else
    echo -e "${GREEN}[OK]${NC} python-docx 已安装"
fi

# ---- 4. 创建 Quick Action workflows ----
echo ""
echo -e "${YELLOW}[3/4]${NC} 创建 Finder 快速操作..."

mkdir -p "$SERVICES_DIR"

# 生成唯一 UUID
generate_uuid() {
    if command -v uuidgen &> /dev/null; then
        uuidgen | tr 'A-F' 'a-f' | tr -d '-'
    else
        python3 -c "import uuid; print(uuid.uuid4().hex)"
    fi
}

# ---- 创建 "用脱敏工具打开" workflow ----
create_workflow() {
    local WF_NAME="$1"
    local WF_DISPLAY_NAME="$2"
    local PY_SCRIPT="$3"
    local WF_DIR="$SERVICES_DIR/$WF_NAME.workflow"
    local WF_CONTENTS="$WF_DIR/Contents"

    mkdir -p "$WF_CONTENTS"

    local UUID1=$(generate_uuid)
    local UUID2=$(generate_uuid)
    local UUID3=$(generate_uuid)
    local UUID4=$(generate_uuid)
    local BUNDLE_ID="com.legal-doc-redactor.$(echo "$WF_NAME" | tr ' ' '-' | tr '[:upper:]' '[:lower:]')"

    # --- document.wflow ---
    cat > "$WF_CONTENTS/document.wflow" << PLISTEOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>actions</key>
    <array>
        <dict>
            <key>action</key>
            <dict>
                <key>AMActionVersion</key>
                <string>1.0.2</string>
                <key>AMApplication</key>
                <array>
                    <string>Automator</string>
                </array>
                <key>AMBundleIdentifier</key>
                <string>com.apple.RunShellScript</string>
                <key>AMName</key>
                <string>Run Shell Script</string>
                <key>AMParameters</key>
                <dict>
                    <key>COMMAND_STRING</key>
                    <string>python3 "$SCRIPT_DIR/$PY_SCRIPT" "\$@"</string>
                    <key>CheckedForUserDefaultShell</key>
                    <true/>
                    <key>inputMethod</key>
                    <integer>1</integer>
                    <key>shell</key>
                    <string>/bin/zsh</string>
                    <key>source</key>
                    <string></string>
                </dict>
                <key>AMProvides</key>
                <dict>
                    <key>Container</key>
                    <string>List</string>
                    <key>Items</key>
                    <array>
                        <string>file paths</string>
                    </array>
                </dict>
                <key>AMRequiredResources</key>
                <array>
                    <dict>
                        <key>AMResourceClass</key>
                        <string>AMFiles</string>
                        <key>AMResourceName</key>
                        <string>Files</string>
                        <key>AMResourceType</key>
                        <string>Files</string>
                    </dict>
                </array>
                <key>AMWarningLevel</key>
                <integer>0</integer>
                <key>ActionBundlePath</key>
                <string>/System/Library/Automator/Run Shell Script.action</string>
                <key>ActionName</key>
                <string>Run Shell Script</string>
                <key>ActionParameters</key>
                <dict>
                    <key>COMMAND_STRING</key>
                    <string>python3 "$SCRIPT_DIR/$PY_SCRIPT" "\$@"</string>
                    <key>CheckedForUserDefaultShell</key>
                    <true/>
                    <key>inputMethod</key>
                    <integer>1</integer>
                    <key>shell</key>
                    <string>/bin/zsh</string>
                    <key>source</key>
                    <string></string>
                </dict>
                <key>BundleIdentifier</key>
                <string>com.apple.RunShellScript</string>
                <key>CFBundleVersion</key>
                <string>1.0.2</string>
                <key>CanShowWhenRun</key>
                <false/>
                <key>Category</key>
                <array>
                    <string>AMCategoryUtility</string>
                </array>
                <key>Class Name</key>
                <string>ShellScriptAction</string>
                <key>InputUUID</key>
                <string>$UUID1</string>
                <key>Keywords</key>
                <array>
                    <string>Shell</string>
                    <string>Script</string>
                    <string>Command</string>
                    <string>Run</string>
                    <string>Execute</string>
                </array>
                <key>OutputUUID</key>
                <string>$UUID2</string>
                <key>UUID</key>
                <string>$UUID3</string>
                <key>UnlocalizedApplications</key>
                <array>
                    <string>Automator</string>
                </array>
                <key>arguments</key>
                <dict>
                    <key>COMMAND_STRING</key>
                    <dict>
                        <key>default</key>
                        <string>python3 "$SCRIPT_DIR/$PY_SCRIPT" "\$@"</string>
                    </dict>
                    <key>CheckedForUserDefaultShell</key>
                    <dict>
                        <key>default</key>
                        <true/>
                        <key>required</key>
                        <string>0</string>
                        <key>type</key>
                        <string>checkbox</string>
                        <key>uuid</key>
                        <string>$UUID4</string>
                    </dict>
                    <key>inputMethod</key>
                    <dict>
                        <key>default</key>
                        <integer>1</integer>
                        <key>required</key>
                        <string>0</string>
                        <key>type</key>
                        <string>popup</string>
                        <key>uuid</key>
                        <string>1</string>
                    </dict>
                    <key>shell</key>
                    <dict>
                        <key>default</key>
                        <string>/bin/zsh</string>
                        <key>required</key>
                        <string>0</string>
                        <key>type</key>
                        <string>popup</string>
                        <key>uuid</key>
                        <string>2</string>
                    </dict>
                    <key>source</key>
                    <dict>
                        <key>default</key>
                        <string></string>
                        <key>required</key>
                        <string>0</string>
                        <key>type</key>
                        <string>text</string>
                        <key>uuid</key>
                        <string>3</string>
                    </dict>
                </dict>
                <key>conversionLabel</key>
                <integer>0</integer>
                <key>disabled</key>
                <false/>
                <key>isView</key>
                <false/>
                <key>isViewVisible</key>
                <true/>
                <key>position</key>
                <dict>
                    <key>x</key>
                    <real>220</real>
                    <key>y</key>
                    <real>20</real>
                </dict>
                <key>resizing</key>
                <dict>
                    <key>allowResizing</key>
                    <true/>
                    <key>expandToAvailableHorizontalSpace</key>
                    <false/>
                    <key>maximumSize</key>
                    <dict/>
                    <key>minimumSize</key>
                    <dict/>
                </dict>
                <key>shouldCollapse</key>
                <false/>
                <key>viewType</key>
                <string>AMAutomatorActionView</string>
            </dict>
            <key>isView</key>
            <false/>
        </dict>
    </array>
    <key>connectors</key>
    <dict/>
    <key>workflowMetaData</key>
    <dict>
        <key>workflowTypeIdentifier</key>
        <string>com.apple.Automator.QuickAction</string>
        <key>inputTypeIdentifier</key>
        <string>com.apple.Automator.filePaths</string>
    </dict>
</dict>
</plist>
PLISTEOF

    # --- Info.plist ---
    cat > "$WF_CONTENTS/Info.plist" << PLISTEOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>$WF_DISPLAY_NAME</string>
    <key>CFBundleIdentifier</key>
    <string>$BUNDLE_ID</string>
    <key>CFBundleVersion</key>
    <string>1.5.0</string>
    <key>NSServices</key>
    <array>
        <dict>
            <key>NSMenuItem</key>
            <dict>
                <key>default</key>
                <string>$WF_DISPLAY_NAME</string>
            </dict>
            <key>NSMessage</key>
            <string>runWorkflowAsService</string>
            <key>NSRequiredContext</key>
            <dict>
                <key>NSApplicationIdentifier</key>
                <string>com.apple.finder</string>
            </dict>
            <key>NSSendFileTypes</key>
            <array>
                <string>org.openxmlformats.wordprocessingml.document</string>
            </array>
        </dict>
    </array>
</dict>
</plist>
PLISTEOF

    echo -e "  ${GREEN}[OK]${NC} $WF_DISPLAY_NAME"
}

create_workflow "用脱敏工具打开" "用脱敏工具打开" "context_menu_server.py"
create_workflow "一键脱敏" "一键脱敏" "quick_redact.py"
create_workflow "一键还原" "一键还原" "quick_restore.py"

# ---- 5. 刷新 Automator 服务缓存 ----
echo ""
echo -e "${YELLOW}[4/4]${NC} 刷新服务缓存..."
/usr/bin/pluginkit --register "$SERVICES_DIR/用脱敏工具打开.workflow" 2>/dev/null || true
/usr/bin/pluginkit --register "$SERVICES_DIR/一键脱敏.workflow" 2>/dev/null || true
/usr/bin/pluginkit --register "$SERVICES_DIR/一键还原.workflow" 2>/dev/null || true
# 触发 Services 菜单刷新
automator --list 2>/dev/null | head -1 > /dev/null || true

echo ""
echo "========================================"
echo "  安装完成！"
echo "========================================"
echo ""
echo "已添加以下右键菜单（快速操作）："
echo "  [1] 用脱敏工具打开 - 在浏览器中打开脱敏工具"
echo "  [2] 一键脱敏       - 直接在后台执行脱敏"
echo "  [3] 一键还原       - 自动查找比对文件并还原"
echo ""
echo "使用方法："
echo "  Finder 中右键点击 .docx 文件 → 快速操作 → 选择对应操作"
echo ""
echo "安装路径: $PROJECT_DIR"
echo ""
echo "卸载方法: 运行 scripts/uninstall_context_menu_mac.sh"
echo ""
