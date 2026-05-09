Attribute VB_Name = "LaTeXTranslator"
Option Explicit

' ============================================================
' LaTeXTranslator - LaTeX → Word线性格式转换模块
' 处理LaTeX和Word线性格式之间的语法差异
' 大部分LaTeX命令(\alpha, \beta, \sum, \int等)可直接通过
' ============================================================

' 主翻译函数：将LaTeX公式内容转为Word兼容的线性格式
Public Function TranslateLaTeX(ByVal latex As String) As String
    Dim result As String
    result = latex

    ' === 顺序敏感: 先处理复杂结构,再处理简单结构 ===

    ' 1. \sqrt[n]{x} → \sqrt(n&x)   (n次方根)
    result = TranslateNthRoot(result)

    ' 2. \sqrt{x} → \sqrt(x)   (平方根)
    result = TranslateSqrt(result)

    ' 3. \begin{cases} ... \end{cases} → \cases( ... )
    result = TranslateEnvironment(result, "cases", "cases")

    ' 4. \begin{pmatrix} ... \end{pmatrix} → \matrix( ... )
    result = TranslateEnvironment(result, "pmatrix", "matrix")

    ' 5. \begin{bmatrix} ... \end{bmatrix} → \matrix( ... )
    result = TranslateEnvironment(result, "bmatrix", "matrix")

    ' 6. \begin{aligned} ... \end{aligned} → \eqarray( ... )
    result = TranslateEnvironment(result, "aligned", "eqarray")

    ' 7. \begin{array} ... \end{array} → \matrix( ... )  (近似)
    result = TranslateEnvironment(result, "array", "matrix")

    ' 8. \left\{ ... \right. → { ... \close
    result = ReplaceLeftRightBrace(result)

    ' 9. \\ → @  (在环境中的换行, 但在普通文本中不应转换)
    ' 此转换在TranslateEnvironment内部完成

    ' 10. 清理不需要的 \left \right \big \biggl \biggr 等定界符修饰
    ' (Word会自动处理括号大小)
    result = CleanDelimiterSizing(result)

    TranslateLaTeX = result
End Function

' ============================================================
' 平方根: \sqrt{...} → \sqrt(...)
' ============================================================

Private Function TranslateSqrt(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' 匹配 \sqrt{...}, 处理嵌套括号
    regex.Pattern = "\\sqrt\{"
    regex.Global = True

    Dim result As String
    result = text

    Dim matches As Object
    Set matches = regex.Execute(text)

    If matches.Count > 0 Then
        ' 需要手动处理, 因为括号需要配对
        result = ReplaceSqrtBraces(result)
    End If

    TranslateSqrt = result
End Function

Private Function ReplaceSqrtBraces(ByVal text As String) As String
    Dim i As Long
    Dim result As String
    result = ""
    i = 1

    Do While i <= Len(text)
        If i + 6 <= Len(text) Then
            If Mid(text, i, 6) = "\sqrt{" Then
                ' 找到配对的闭合 }
                Dim braceCount As Long
                Dim j As Long
                braceCount = 1
                j = i + 6

                Do While j <= Len(text) And braceCount > 0
                    If Mid(text, j, 1) = "{" Then
                        braceCount = braceCount + 1
                    ElseIf Mid(text, j, 1) = "}" Then
                        braceCount = braceCount - 1
                    End If

                    If braceCount > 0 Then
                        j = j + 1
                    End If
                Loop

                If braceCount = 0 Then
                    ' i+6 是 { 后第一个字符, j-1 是配对的 }
                    Dim innerContent As String
                    innerContent = Mid(text, i + 6, j - i - 6)

                    result = result & "\sqrt(" & innerContent & ")"
                    i = j + 1
                Else
                    ' 未找到配对,保留原文
                    result = result & "\sqrt{"
                    i = i + 6
                End If
            Else
                result = result & Mid(text, i, 1)
                i = i + 1
            End If
        Else
            result = result & Mid(text, i, 1)
            i = i + 1
        End If
    Loop

    ReplaceSqrtBraces = result
End Function

' ============================================================
' n次方根: \sqrt[n]{x} → \sqrt(n&x)
' ============================================================

Private Function TranslateNthRoot(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "\\sqrt\[([^\]]+)\]\{"
    regex.Global = True

    Dim result As String
    result = regex.Replace(text, "\sqrt($1&")

    ' 现在处理 sqrt( n & ... { 中的 { 转 )
    ' 需要找到 & 后的 { 并转为 )
    ' 这里先简化: 替换 & 后的 { 为 ( 并在内容末尾加 )
    ' 实际上需要括号配对,走手动处理

    If InStr(result, "\sqrt(") > 0 And InStr(result, "&") > 0 Then
        result = FixNthRootBraces(result)
    End If

    TranslateNthRoot = result
End Function

Private Function FixNthRootBraces(ByVal text As String) As String
    ' 查找 \sqrt(n&{...}) 模式,将{...}转为(...)
    Dim i As Long
    Dim result As String
    result = ""
    i = 1

    Do While i <= Len(text)
        If i + 6 <= Len(text) Then
            If Mid(text, i, 6) = "\sqrt(" Then
                ' 查找 & 符号
                Dim ampPos As Long
                ampPos = InStr(i, text, "&")

                If ampPos > 0 And ampPos < i + 50 Then
                    ' & 后面应该是 {...}
                    Dim afterAmp As String
                    afterAmp = Mid(text, ampPos + 1)

                    If Len(afterAmp) > 0 And Left(afterAmp, 1) = "{" Then
                        ' 找到配对的 }
                        Dim braceCount2 As Long
                        Dim k As Long
                        braceCount2 = 1
                        k = ampPos + 2

                        Do While k <= Len(text) And braceCount2 > 0
                            If Mid(text, k, 1) = "{" Then
                                braceCount2 = braceCount2 + 1
                            ElseIf Mid(text, k, 1) = "}" Then
                                braceCount2 = braceCount2 - 1
                            End If
                            If braceCount2 > 0 Then k = k + 1
                        Loop

                        If braceCount2 = 0 Then
                            Dim nthContent As String
                            nthContent = Mid(text, ampPos + 2, k - ampPos - 2)

                            ' 找到 sqrt 的前半部分 ( )
                            Dim exp As String
                            exp = Mid(text, i + 6, ampPos - i - 6) ' n的值

                            result = result & "\sqrt(" & exp & "&" & nthContent & ")"
                            i = k + 1
                        Else
                            result = result & Mid(text, i, 1)
                            i = i + 1
                        End If
                    Else
                        result = result & Mid(text, i, 1)
                        i = i + 1
                    End If
                Else
                    result = result & Mid(text, i, 1)
                    i = i + 1
                End If
            Else
                result = result & Mid(text, i, 1)
                i = i + 1
            End If
        Else
            result = result & Mid(text, i, 1)
            i = i + 1
        End If
    Loop

    FixNthRootBraces = result
End Function

' ============================================================
' 环境翻译: \begin{env}...\end{env} → \func(...)
' ============================================================

Private Function TranslateEnvironment(ByVal text As String, ByVal envName As String, ByVal wordFunc As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "\\begin\{" & envName & "\}([\s\S]*?)\\end\{" & envName & "\}"
    regex.Global = True

    If Not regex.Test(text) Then
        TranslateEnvironment = text
        Exit Function
    End If

    Dim result As String
    result = text

    Dim matches As Object
    Set matches = regex.Execute(text)

    Dim m As Object
    For Each m In matches
        Dim innerContent As String
        innerContent = m.SubMatches(0)

        ' 将 \\ 替换为 @ (矩阵行分隔符)
        innerContent = Replace(innerContent, "\\", "@")

        ' 将 & 保留 (列分隔符, Word也用&)

        ' 构造Word格式
        Dim replacement As String
        replacement = "\" & wordFunc & "(" & innerContent & ")"

        result = Replace(result, m.Value, replacement)
    Next m

    TranslateEnvironment = result
End Function

' ============================================================
' 左大括号: \left\{ ... \right. → { ... \close
' ============================================================

Private Function ReplaceLeftRightBrace(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "\\left\\\{([\s\S]*?)\\right\."
    regex.Global = True

    ReplaceLeftRightBrace = regex.Replace(text, "{$1\close")
End Function

' ============================================================
' 清理不必要的定界符修饰 (\left \right \big等)
' Word会自动处理括号大小, 但保留\left\right让BuildUp处理
' ============================================================

Private Function CleanDelimiterSizing(ByVal text As String) As String
    ' 目前不做清理, Word's BuildUp 会处理这些命令
    ' 如果发现转换后有多余文本,再在此处理
    CleanDelimiterSizing = text
End Function
