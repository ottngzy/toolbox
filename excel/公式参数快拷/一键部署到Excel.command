#!/bin/bash
# ============================================================
# 公式参数快拷 —— 一键部署到 Excel
# 双击运行即可：自动打开 Excel、打开宏编辑器、新建模块并粘贴代码
# ============================================================
cd "$(dirname "$0")" || exit 1

echo "════════════════════════════════════════"
echo "  公式参数快拷 · 一键部署到 Excel"
echo "════════════════════════════════════════"

# 1. 把 VBA 代码放进剪贴板（去掉 .bas 文件的属性行，可直接粘贴）
if [ -f "公式参数快拷.bas" ]; then
    tail -n +2 "公式参数快拷.bas" | pbcopy
    echo "✔ VBA 代码已复制到剪贴板"
else
    echo "✘ 找不到 公式参数快拷.bas，请确认它和本安装器在同一文件夹"
    read -r -p "按回车退出..."
    exit 1
fi

# 2. 打开 Excel
echo "✔ 正在打开 Excel..."
open -a "Microsoft Excel"
sleep 4

# 3. 通过图形界面自动化：打开 VBE、插入模块、粘贴代码
echo "✔ 正在自动打开宏编辑器并粘贴代码（如首次运行，macOS 可能要求授权）..."

osascript <<'APPLESCRIPT' 2>/dev/null
tell application "Microsoft Excel" to activate
delay 2

tell application "System Events"
    -- 打开 VBE：工具 > 宏 > Visual Basic 编辑器（中英文菜单都试）
    set openedVBE to false
    try
        click menu item "Visual Basic 编辑器" of menu of menu item "宏" of menu "工具" of menu bar 1 of process "Microsoft Excel"
        set openedVBE to true
    end try
    if not openedVBE then
        try
            click menu item "Visual Basic Editor" of menu of menu item "Macro" of menu "Tools" of menu bar 1 of process "Microsoft Excel"
            set openedVBE to true
        end try
    end if
end tell
delay 3

tell application "System Events"
    -- 插入 > 模块
    set inserted to false
    repeat with procName in {"Microsoft Excel", "Microsoft Visual Basic"}
        try
            tell process procName
                try
                    click menu item "模块" of menu "插入" of menu bar 1
                    set inserted to true
                on error
                    click menu item "Module" of menu "Insert" of menu bar 1
                    set inserted to true
                end try
            end tell
        end try
        if inserted then exit repeat
    end repeat
    delay 1
    -- 粘贴剪贴板里的代码
    keystroke "v" using command down
end tell
APPLESCRIPT

if [ $? -eq 0 ]; then
    echo ""
    echo "════════════════════════════════════════"
    echo "  ✅ 代码应该已经粘贴进宏编辑器了"
    echo "════════════════════════════════════════"
else
    echo ""
    echo "════════════════════════════════════════"
    echo "  ⚠️  自动操作被 macOS 拦截（或 Excel 界面未就绪）"
    echo "════════════════════════════════════════"
fi

cat <<'TIPS'

接下来只需 2 步：

  ① 看一眼 Excel 的宏编辑器：如果右边代码窗口里出现了
     「Sub 公式参数快拷()」，说明成功了；如果是空的，
     在编辑器里手动按 Cmd+V 粘贴一下即可
     （代码已在你剪贴板里）

  ② 以后使用：选中含公式的单元格 → 按 Option+F8
     → 选「公式参数快拷」→ 运行

如果第 ① 步是空的且 Cmd+V 也不行：
  · macOS 需要在「系统设置 → 隐私与安全性 → 辅助功能」里
    给「终端」打勾，然后重新双击本安装器
  · 或手动操作：Excel 菜单「工具 → 宏 → Visual Basic 编辑器」
    →「插入 → 模块」→ Cmd+V 粘贴

TIPS

read -r -p "按回车退出..."
