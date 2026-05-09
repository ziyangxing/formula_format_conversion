Attribute VB_Name = "EquationBuilder"
Option Explicit

' ============================================================
' EquationBuilder - OMath公式对象创建模块
' 在指定位置创建Word公式对象并调用BuildUp转换为专业格式
' ============================================================

' 构建公式：将内容转为Word专业公式
Public Sub BuildEquation(ByVal doc As Document, ByVal match As FormulaMatch, ByVal translatedContent As String)
    On Error GoTo ErrorHandler

    Dim formulaRange As Range

    ' 获取公式在文档中的绝对位置范围
    Set formulaRange = doc.Range(match.StartPos, match.EndPos + 1)

    ' 检查内容是否为空
    If Len(Trim(translatedContent)) = 0 Then
        Exit Sub
    End If

    ' 将范围内的文本替换为公式内容（不含分隔符）
    formulaRange.Text = translatedContent

    ' 重新调整 Range 为实际内容
    Set formulaRange = doc.Range(match.StartPos, match.StartPos + Len(translatedContent))

    ' 创建 OMath 公式对象
    Dim objOMath As OMath
    Set objOMath = doc.OMaths.Add(formulaRange)

    ' 设置公式类型：行内 or 显示
    If match.IsDisplay Then
        objOMath.Type = wdOMathDisplay
    Else
        objOMath.Type = wdOMathInline
    End If

    ' 调用 BuildUp 转换为专业格式
    objOMath.BuildUp

    Exit Sub

ErrorHandler:
    ' 如果创建公式失败(如API不支持),保持原文不变
    ' 静默失败,继续处理下一个公式
    Debug.Print "EquationBuilder error at position " & match.StartPos & ": " & Err.Description
End Sub

' 获取文档中所有公式区域（用于调试/批量操作）
Public Function GetEquationRanges(ByVal doc As Document) As Collection
    Dim result As New Collection
    Dim i As Long

    For i = 1 To doc.OMaths.Count
        result.Add doc.OMaths(i).Range
    Next i

    Set GetEquationRanges = result
End Function

' 检查指定位置是否已有公式对象
Public Function HasExistingEquation(ByVal doc As Document, ByVal startPos As Long, ByVal endPos As Long) As Boolean
    Dim i As Long
    Dim eqRange As Range

    For i = 1 To doc.OMaths.Count
        Set eqRange = doc.OMaths(i).Range
        ' 检查是否有重叠
        If startPos <= eqRange.End And endPos >= eqRange.Start Then
            HasExistingEquation = True
            Exit Function
        End If
    Next i
End Function
