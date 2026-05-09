Attribute VB_Name = "FormulaValidator"
Option Explicit

' ============================================================
' FormulaValidator - 公式验证模块
' 过滤假阳性匹配：货币金额、代码块内、转义符、非数学内容
' match.StartPos/EndPos 为文档绝对位置 (0-based)
' docRange.Text 为选区文本 (1-based)
' ============================================================

' 综合验证入口：判断一个公式匹配是否有效
Public Function IsValidFormula(ByVal docRange As Range, ByVal match As FormulaMatch) As Boolean
    ' 计算 match 在 docRange.Text 中的相对位置 (1-based)
    Dim textRelStart As Long
    textRelStart = match.StartPos - docRange.Start + 1

    If IsCurrencyDollar(match) Then Exit Function
    If IsEscapedDelimiter(docRange, textRelStart) Then Exit Function
    If IsInsideCodeBlock(docRange, textRelStart) Then Exit Function
    If Not LooksLikeMath(match.Content) Then Exit Function

    IsValidFormula = True
End Function

' ============================================================
' 验证1: 货币金额检测 ($后紧跟数字且无数学符号)
' ============================================================

Private Function IsCurrencyDollar(ByVal match As FormulaMatch) As Boolean
    ' 只对 $ 分隔符进行货币检测
    If match.DelimiterType <> "$" Then Exit Function

    Dim content As String
    content = Trim(match.Content)

    If Len(content) = 0 Then
        IsCurrencyDollar = True
        Exit Function
    End If

    ' 如果内容以数字开头,检查是否包含数学符号
    If IsNumeric(Left(content, 1)) Or Left(content, 1) = "." Or Left(content, 1) = "," Then
        ' 以数字/小数点/逗号开始 → 疑似货币
        If Not ContainsMathSymbols(content) Then
            IsCurrencyDollar = True
        End If
    End If
End Function

' ============================================================
' 验证2: 转义符检测 (\$ 不作为分隔符)
' ============================================================

Private Function IsEscapedDelimiter(ByVal docRange As Range, ByVal textRelStart As Long) As Boolean
    ' textRelStart: match 在 docRange.Text 中的起始位置 (1-based)
    ' 确保 match 不在选区最开头
    If textRelStart <= 1 Then Exit Function
    ' 确保 range 不是从文档最开头开始的 (避免向前溢出)
    If docRange.Start <= 0 Then Exit Function

    Dim precedingChar As String
    precedingChar = Mid(docRange.Text, textRelStart - 1, 1)

    If precedingChar = "\" Then
        ' 再确认前面不是第二个反斜杠 (\\$ 不是转义)
        If textRelStart <= 2 Or Mid(docRange.Text, textRelStart - 2, 1) <> "\" Then
            IsEscapedDelimiter = True
        End If
    End If
End Function

' ============================================================
' 验证3: 代码块检测 (围栏、行内)
' ============================================================

Private Function IsInsideCodeBlock(ByVal docRange As Range, ByVal textRelStart As Long) As Boolean
    If IsInsideFencedCodeBlock(docRange, textRelStart) Then
        IsInsideCodeBlock = True
        Exit Function
    End If

    If IsInsideInlineCode(docRange, textRelStart) Then
        IsInsideCodeBlock = True
        Exit Function
    End If
End Function

' 检测围栏代码块 ``` ```
Private Function IsInsideFencedCodeBlock(ByVal docRange As Range, ByVal textRelStart As Long) As Boolean
    Dim text As String
    Dim textBefore As String
    Dim fenceCount As Long
    Dim pos As Long

    text = docRange.Text

    ' textRelStart <= 1 → match 在选区最开头，没有 preceding text
    If textRelStart <= 1 Then Exit Function

    textBefore = Left(text, textRelStart - 1)

    ' 统计 match 之前的 ``` 出现次数
    pos = 1
    Do While pos <= Len(textBefore)
        Dim foundPos As Long
        foundPos = InStr(pos, textBefore, "```")
        If foundPos = 0 Then Exit Do

        fenceCount = fenceCount + 1
        pos = foundPos + 3
    Loop

    ' 奇数个 ``` → 在代码块内
    If fenceCount Mod 2 = 1 Then
        IsInsideFencedCodeBlock = True
    End If
End Function

' 检测行内代码 `...`
Private Function IsInsideInlineCode(ByVal docRange As Range, ByVal textRelStart As Long) As Boolean
    Dim text As String
    Dim textLine As String
    Dim backtickCount As Long
    Dim i As Long

    text = docRange.Text

    ' 获取 match 所在行的内容（从上一个换行到下一个换行）
    Dim lineStart As Long
    Dim lineEnd As Long

    lineStart = textRelStart
    Do While lineStart > 1
        Dim c As String
        c = Mid(text, lineStart - 1, 1)
        If c = vbCr Or c = vbLf Then Exit Do
        lineStart = lineStart - 1
    Loop

    lineEnd = textRelStart
    Do While lineEnd <= Len(text)
        Dim c2 As String
        c2 = Mid(text, lineEnd, 1)
        If c2 = vbCr Or c2 = vbLf Then Exit Do
        lineEnd = lineEnd + 1
    Loop

    textLine = Mid(text, lineStart, lineEnd - lineStart)

    ' 统计 match 之前反引号数量
    Dim relPos As Long
    relPos = textRelStart - lineStart + 1

    For i = 1 To relPos - 1
        If Mid(textLine, i, 1) = "`" Then
            backtickCount = backtickCount + 1
        End If
    Next i

    ' 奇数个反引号 → 在行内代码中
    If backtickCount Mod 2 = 1 Then
        IsInsideInlineCode = True
    End If
End Function

' ============================================================
' 验证4: 数学特征检测 (内容是否"像数学公式")
' ============================================================

Private Function LooksLikeMath(ByVal content As String) As Boolean
    ' 空内容不转换
    If Len(Trim(content)) = 0 Then Exit Function

    ' 检查是否包含至少一个数学特征
    If ContainsMathSymbols(content) Then
        LooksLikeMath = True
        Exit Function
    End If

    ' 不包含任何数学符号,但可能是简单表达式如 x^2 或 a+b=c
    ' 通过前面的货币检测后,大概率是真公式
    LooksLikeMath = True
End Function

' 检查字符串中是否包含数学符号
Public Function ContainsMathSymbols(ByVal text As String) As Boolean
    ' LaTeX命令
    If InStr(text, "\") > 0 Then
        ContainsMathSymbols = True
        Exit Function
    End If

    ' 上下标
    If InStr(text, "^") > 0 Or InStr(text, "_") > 0 Then
        ContainsMathSymbols = True
        Exit Function
    End If

    ' 花括号（LaTeX分组）
    If InStr(text, "{") > 0 Or InStr(text, "}") > 0 Then
        ContainsMathSymbols = True
        Exit Function
    End If

    ' 数学运算符
    Dim ops As Variant
    ops = Array("+", "-", "*", "=", "<", ">", "/", "|", "!")

    Dim v As Variant
    For Each v In ops
        If InStr(text, CStr(v)) > 0 Then
            ContainsMathSymbols = True
            Exit Function
        End If
    Next v

    ' 希腊字母 Unicode
    Dim i As Long
    For i = 1 To Len(text)
        Dim ascVal As Long
        ascVal = AscW(Mid(text, i, 1))
        ' 希腊字母范围: U+0391-U+03C9
        If ascVal >= &H391 And ascVal <= &H3C9 Then
            ContainsMathSymbols = True
            Exit Function
        End If
    Next i
End Function
