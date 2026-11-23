Attribute VB_Name = "FormulaParamCopy"
Option Explicit

' ============================================================
' 公式参数快拷
' 用法：选中一个或多个含公式的单元格（如 A1:C1），运行本宏
' 效果：逐格提取参数，按指定方向（下/右/左/上）批量生成同参数公式
' ============================================================
Sub 公式参数快拷()

    Dim sel As Range, c As Range
    Dim src As String, arg As String, previewArg As String
    Dim p1 As Long, p2 As Long
    Dim inputFns As String, dir As String
    Dim fns As Variant, f As Variant
    Dim i As Long, n As Long

    ' 1. 检查选区
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中一个或多个含公式的单元格再运行。", vbExclamation
        Exit Sub
    End If
    Set sel = Selection

    ' 2. 找到第一个含公式且带括号的单元格，用于预览参数
    For Each c In sel.Cells
        If c.HasFormula Then
            src = c.Formula
            p1 = InStr(src, "(")
            p2 = InStrRev(src, ")")
            If p1 > 0 And p2 > p1 Then
                previewArg = Mid(src, p1 + 1, p2 - p1 - 1)
                Exit For
            End If
        End If
    Next c
    If previewArg = "" Then
        MsgBox "选区内没有找到含 (参数) 的公式。", vbExclamation
        Exit Sub
    End If

    ' 3. 询问要生成哪些函数
    inputFns = InputBox( _
        "将对选区内每个含公式的单元格分别处理。" & vbCrLf & _
        "第一个公式的参数是：" & previewArg & vbCrLf & vbCrLf & _
        "请输入要生成的函数（英文逗号分隔）：", _
        "公式参数快拷", "MEDIAN, MIN, MAX")
    If inputFns = "" Then Exit Sub

    ' 4. 询问生成方向
    dir = InputBox("公式往哪个方向生成？" & vbCrLf & _
                   "输入：下 / 右 / 左 / 上", _
                   "公式参数快拷", "下")
    If dir = "" Then Exit Sub

    fns = Split(inputFns, ",")

    ' 5. 逐格提取参数、按方向写入公式
    Application.ScreenUpdating = False
    n = 0
    For Each c In sel.Cells
        If c.HasFormula Then
            src = c.Formula
            p1 = InStr(src, "(")
            p2 = InStrRev(src, ")")
            If p1 > 0 And p2 > p1 Then
                arg = Mid(src, p1 + 1, p2 - p1 - 1)
                i = 1
                For Each f In fns
                    f = Trim(CStr(f))
                    If f <> "" Then
                        Select Case dir
                            Case "右"
                                c.Offset(0, i).Formula = "=" & f & "(" & arg & ")"
                            Case "左"
                                c.Offset(0, -i).Formula = "=" & f & "(" & arg & ")"
                            Case "上"
                                c.Offset(-i, 0).Formula = "=" & f & "(" & arg & ")"
                            Case Else
                                c.Offset(i, 0).Formula = "=" & f & "(" & arg & ")"
                        End Select
                        i = i + 1
                    End If
                Next f
                n = n + 1
            End If
        End If
    Next c
    Application.ScreenUpdating = True

    MsgBox "完成！共处理了 " & n & " 个公式单元格。", vbInformation

End Sub
