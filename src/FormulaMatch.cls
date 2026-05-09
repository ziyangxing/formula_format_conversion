VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaMatch"
Option Explicit

' ============================================================
' FormulaMatch - 公式匹配信息类
' 存储单个公式匹配的位置、内容和类型
' ============================================================

Public StartPos As Long
Public EndPos As Long
Public Content As String
Public IsDisplay As Boolean
Public DelimiterType As String    ' "$", "$$", "\(", "\["

Public Function MatchLength() As Long
    MatchLength = EndPos - StartPos + 1
End Function
