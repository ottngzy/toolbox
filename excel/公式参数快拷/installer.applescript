-- ============================================================
-- 公式参数快拷 · 一键安装器（覆盖更新版）
-- 双击运行：复制最新 VBA 代码 → 打开 Excel → 打开宏编辑器 → 全选旧代码 → 粘贴覆盖
-- ============================================================
on run
	-- 1. 读取随 App 打包的 .bas 文件，去掉属性行后放入剪贴板
	try
		set resPath to (path to resource "公式参数快拷.bas")
		set fRef to open for access resPath
		set fileContent to read fRef as «class utf8»
		close access fRef
	on error
		display dialog "⚠️ 找不到内置的 VBA 代码文件，请重新下载安装器。" buttons {"退出"} default button 1 with title "公式参数快拷安装器"
		return
	end try

	set saveTID to AppleScript's text item delimiters
	set AppleScript's text item delimiters to return
	set allLines to text items of fileContent
	if (count of allLines) > 1 then
		set codeText to (items 2 thru -1 of allLines) as string
	else
		set codeText to fileContent
	end if
	set AppleScript's text item delimiters to saveTID
	set the clipboard to codeText

	-- 2. 打开 Excel
	try
		tell application "Microsoft Excel" to activate
	on error
		display dialog "⚠️ 没检测到 Microsoft Excel，请先安装 Excel 再运行本安装器。" buttons {"退出"} default button 1 with title "公式参数快拷安装器"
		return
	end try
	delay 2

	-- 3. 自动打开 VBE（工具 > 宏 > Visual Basic 编辑器，中英文菜单都试）
	set openedVBE to false
	tell application "System Events"
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

	-- 4. 覆盖粘贴：VBE 打开时会恢复上次编辑的模块，直接全选替换
	set pasted to false
	if openedVBE then
		tell application "System Events"
			try
				keystroke "a" using command down
				delay 0.5
				keystroke "v" using command down
				set pasted to true
			end try
		end tell
	end if
	delay 1

	-- 5. 检测代码是否真的进了编辑窗格（尽力读取窗格文本）
	set paneOK to false
	if pasted then
		try
			tell application "System Events"
				tell process "Microsoft Excel"
					set paneText to ""
					try
						set paneText to value of text area 1 of scroll area 1 of window 1
					end try
					if paneText does not contain "公式参数快拷" then
						try
							set paneText to value of text area 1 of window 1
						end try
					end if
				end tell
			tell
			if paneText contains "公式参数快拷" then set paneOK to true
		end try
	end if

	-- 6. 汇报结果
	if pasted and paneOK then
		display dialog "✅ 已覆盖更新！" & return & return & "最新 VBA 代码已粘贴进宏编辑器（替换了旧代码）。" & return & return & "使用方法：选中含公式的单元格区域 → 按 Option+F8 → 运行「公式参数快拷」，可自定义函数列表和生成方向（下/右/左/上）。" buttons {"完成"} default button 1 with title "公式参数快拷安装器"
	else if pasted then
		display dialog "⚠️ 已执行粘贴，但无法确认结果" & return & return & "最常见的原因是宏编辑器里没有打开任何模块（灰色空白）。请手动完成：" & return & "在编辑器左侧工程窗口里双击你原来的模块（如 Module1）→ 点进右边代码窗口 → 按 Cmd+A 全选 → Cmd+V 粘贴（代码已复制好，直接粘即可）。" buttons {"好的"} default button 1 with title "公式参数快拷安装器"
	else
		display dialog "⚠️ 自动操作没有完成" & return & return & "VBA 代码已经复制到你的剪贴板，请手动完成：" & return & "Excel 菜单「工具 → 宏 → Visual Basic 编辑器」→ 双击原来的模块 → 代码窗口里 Cmd+A 全选 → Cmd+V 粘贴覆盖。" & return & return & "如果希望以后全自动：系统设置 → 隐私与安全性 → 辅助功能，把「公式参数快拷安装器」打勾，然后重新双击本程序。" buttons {"好的"} default button 1 with title "公式参数快拷安装器"
	end if
end run
