-- ============================================================
-- 公式参数快拷 · 一键安装器
-- 双击运行：自动复制 VBA 代码 → 打开 Excel → 打开宏编辑器 → 插入模块 → 粘贴
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

	-- 4. 插入 > 模块，然后粘贴
	set inserted to false
	if openedVBE then
		tell application "System Events"
			try
				try
					click menu item "模块" of menu "插入" of menu bar 1 of process "Microsoft Excel"
				on error
					click menu item "Module" of menu "Insert" of menu bar 1 of process "Microsoft Excel"
				end try
				set inserted to true
			end try
		end tell
	end if
	if inserted then
		delay 1
		tell application "System Events" to keystroke "v" using command down
	end if

	-- 5. 汇报结果
	if openedVBE and inserted then
		display dialog "✅ 部署完成！" & return & return & "VBA 代码已自动粘贴进宏编辑器。" & return & return & "使用方法：选中含公式的单元格 → 按 Option+F8 → 运行「公式参数快拷」，会依次问你函数列表和生成方向（下/右/左/上）。" buttons {"完成"} default button 1 with title "公式参数快拷安装器"
	else
		display dialog "⚠️ 自动操作没有完成" & return & return & "VBA 代码已经复制到你的剪贴板，请手动完成最后一步：" & return & "Excel 菜单「工具 → 宏 → Visual Basic 编辑器」→「插入 → 模块」→ 按 Cmd+V 粘贴。" & return & return & "如果希望以后全自动：系统设置 → 隐私与安全性 → 辅助功能，把「公式参数快拷安装器」打勾，然后重新双击本程序。" buttons {"好的"} default button 1 with title "公式参数快拷安装器"
	end if
end run
