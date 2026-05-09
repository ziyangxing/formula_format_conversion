Attribute VB_Name = "FormulaConverter"
Option Explicit

' ============================================================
' FormulaConverter - 主入口模块
' 协调查找→验证→翻译→构建的完整流程
' 暴露给Word Ribbon按钮调用
' ============================================================

' 主入口：转换选区内的所有公式
Public Sub ConvertFormulasInSelection()
    Dim doc As Document
    Set doc = ActiveDocument

    ' 安全检查：确保有选区
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选择要转换的文本区域（Ctrl+A 全选）", vbInformation, "公式转换"
        Exit Sub
    End If

    ' 关闭屏幕更新以提升性能
    Application.ScreenUpdating = False

    Dim selRange As Range
    Set selRange = Selection.Range

    ' === 步骤1: 扫描公式 ===
    Dim formulas As Collection
    Set formulas = FindAllFormulas(selRange)

    If formulas.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "未检测到公式。支持的格式：" & vbCrLf & _
               "  $...$  (行内公式)" & vbCrLf & _
               "  $$...$$  (显示公式)" & vbCrLf & _
               "  \(...\)  (行内公式)" & vbCrLf & _
               "  \[...\]  (显示公式)", vbInformation, "公式转换"
        Exit Sub
    End If

    ' === 步骤2-6: 处理每个公式 (从后往前) ===
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim skippedReasons As String

    Dim match As FormulaMatch
    For Each match In formulas

        ' 验证
        If Not IsValidFormula(selRange, match) Then
            skippedCount = skippedCount + 1
            GoTo NextMatch
        End If

        ' LaTeX → Word线性格式翻译
        Dim translated As String
        translated = TranslateLaTeX(match.Content)

        ' 构建公式
        BuildEquation doc, match, translated

        convertedCount = convertedCount + 1

NextMatch:
    Next match

    ' 恢复屏幕更新
    Application.ScreenUpdating = True

    ' 显示结果
    Dim msg As String
    msg = "转换完成！" & vbCrLf & vbCrLf & _
          "成功转换: " & convertedCount & " 个公式"

    If skippedCount > 0 Then
        msg = msg & vbCrLf & "跳过: " & skippedCount & " 个（非公式内容或代码块内）"
    End If

    MsgBox msg, vbInformation, "公式转换"
End Sub

' 转换整个文档（不依赖选区）
Public Sub ConvertFormulasInDocument()
    Dim doc As Document
    Set doc = ActiveDocument

    ' 保存当前选区
    Dim originalRange As Range
    Set originalRange = Selection.Range

    ' 全选
    Selection.WholeStory

    ' 执行转换
    ConvertFormulasInSelection

    ' 恢复选区
    originalRange.Select
End Sub
