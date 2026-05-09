Attribute VB_Name = "FormulaFinder"
Option Explicit

' ============================================================
' FormulaFinder - 公式分隔符扫描模块
' 采用逐字符遍历，避免正则的边界歧义（$ vs $$）
' 返回 Collection of FormulaMatch 对象，按文档位置从后往前排序
' 匹配的 StartPos/EndPos 为文档绝对位置 (0-based)
' ============================================================

' 扫描文档范围，找到所有公式匹配项
' 返回 Collection of FormulaMatch 对象（从后往前排序）
Public Function FindAllFormulas(ByVal docRange As Range) As Collection
    Dim results As New Collection
    Dim tempResults As New Collection
    Dim text As String
    Dim i As Long
    Dim textLen As Long
    Dim baseOffset As Long

    text = docRange.Text
    textLen = Len(text)

    ' 文本位置 → 文档绝对位置的偏移量
    ' 文本position 1 对应文档 position docRange.Start
    baseOffset = docRange.Start

    If textLen < 2 Then
        Set FindAllFormulas = results
        Exit Function
    End If

    i = 1
    Do While i < textLen

        ' --- 检测 $$ 显示公式 ---
        If i < textLen And Mid(text, i, 2) = "$$" Then
            TryMatchDisplayDollar text, textLen, i, tempResults, baseOffset

        ' --- 检测 $ 行内公式 (但不能是$$, 不能是转义\$ ) ---
        ElseIf Mid(text, i, 1) = "$" Then
            ' 确保不是 $$ 的第二个字符
            If i = 1 Or Mid(text, i - 1, 1) <> "$" Then
                ' 确保不是 转义\$
                If i = 1 Or Mid(text, i - 1, 1) <> "\" Then
                    TryMatchInlineDollar text, textLen, i, tempResults, baseOffset
                End If
            End If

        ' --- 检测 \[ 显示公式 ---
        ElseIf i < textLen And Mid(text, i, 2) = "\[" Then
            TryMatchBracket text, textLen, i, tempResults, "\[", "\]", True, baseOffset

        ' --- 检测 \( 行内公式 ---
        ElseIf i < textLen And Mid(text, i, 2) = "\(" Then
            TryMatchBracket text, textLen, i, tempResults, "\(", "\)", False, baseOffset

        End If

        i = i + 1
    Loop

    ' 反向复制到 results（从后往前排序）
    Dim j As Long
    For j = tempResults.Count To 1 Step -1
        results.Add tempResults(j)
    Next j

    Set FindAllFormulas = results
End Function

' ============================================================
' 辅助匹配函数
' ============================================================

Private Sub TryMatchDisplayDollar(ByVal text As String, ByVal textLen As Long, ByRef i As Long, ByRef results As Collection, ByVal baseOffset As Long)
    Dim closePos As Long
    closePos = FindClosingDelimiter(text, i + 2, textLen, "$$")

    If closePos > 0 Then
        AddMatch results, i, closePos + 1, text, True, "$$", baseOffset
        i = closePos + 1
    End If
End Sub

Private Sub TryMatchInlineDollar(ByVal text As String, ByVal textLen As Long, ByRef i As Long, ByRef results As Collection, ByVal baseOffset As Long)
    Dim j As Long
    j = i + 1

    Do While j <= textLen
        Dim ch As String
        ch = Mid(text, j, 1)

        ' 换行符 - 行内公式不能跨行
        If ch = vbCr Or ch = vbLf Or ch = vbVerticalTab Then
            Exit Do
        End If

        ' 找到闭合 $,且不是转义\$
        If ch = "$" Then
            If Mid(text, j - 1, 1) <> "\" Then
                ' 确保不是 $$ (如果是$$, 这里应该是开头匹配的$$)
                If j >= textLen Or Mid(text, j + 1, 1) <> "$" Then
                    AddMatch results, i, j, text, False, "$", baseOffset
                    i = j
                    Exit Sub
                End If
            End If
        End If

        j = j + 1
    Loop
End Sub

Private Sub TryMatchBracket(ByVal text As String, ByVal textLen As Long, ByRef i As Long, ByRef results As Collection, ByVal openDelim As String, ByVal closeDelim As String, ByVal isDisplay As Boolean, ByVal baseOffset As Long)
    Dim closePos As Long
    closePos = FindClosingDelimiter(text, i + Len(openDelim), textLen, closeDelim)

    If closePos > 0 Then
        AddMatch results, i, closePos + Len(closeDelim) - 1, text, isDisplay, openDelim, baseOffset
        i = closePos + Len(closeDelim) - 1
    End If
End Sub

' 在 text 中从 startFrom 开始查找 delimiter
Private Function FindClosingDelimiter(ByVal text As String, ByVal startFrom As Long, ByVal textLen As Long, ByVal delimiter As String) As Long
    Dim delimLen As Long
    delimLen = Len(delimiter)

    If startFrom + delimLen - 1 > textLen Then
        FindClosingDelimiter = 0
        Exit Function
    End If

    Dim pos As Long
    For pos = startFrom To textLen - delimLen + 1
        If Mid(text, pos, delimLen) = delimiter Then
            FindClosingDelimiter = pos
            Exit Function
        End If
    Next pos

    FindClosingDelimiter = 0
End Function

' 添加匹配到集合
' startPos/endPos 是文本内偏移(1-based)，baseOffset = docRange.Start 转为文档绝对位置
Private Sub AddMatch(ByRef results As Collection, ByVal startPos As Long, ByVal endPos As Long, ByVal text As String, ByVal isDisplay As Boolean, ByVal delimType As String, ByVal baseOffset As Long)
    Dim match As FormulaMatch
    Set match = New FormulaMatch
    Dim delimLen As Long

    delimLen = Len(delimType)
    ' 转为文档绝对位置: text position 1 → document position baseOffset
    match.StartPos = startPos + baseOffset - 1
    match.EndPos = endPos + baseOffset - 1
    match.Content = Mid(text, startPos + delimLen, endPos - startPos - delimLen + 1 - delimLen)
    match.IsDisplay = isDisplay
    match.DelimiterType = delimType

    results.Add match
End Sub

