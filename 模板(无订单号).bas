Attribute VB_Name = "无订单号快递"
Private Sub 删除()
    Dim rng As Range
    Dim i As Integer
    jrng = Range("c65536").End(xlUp).Row
        For Each rng In Range("c1:c" & jrng)
            If rng.Value = "x1" Then
            rng.EntireRow.Select
            Selection.Delete Shift:=xlUp
            End If
        Next
        Range("a1").EntireRow.Delete
End Sub
Private Sub 分列()
Application.DisplayAlerts = False
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
       :=Array(Array(2, 1)), TrailingMinusNumbers:=True
        Application.DisplayAlerts = True
End Sub

Private Sub 合并()
    jrng = Range("c65536").End(xlUp).Row
    Range("L1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-5],RC[-4],RC[-3],RC[-2])"
    Range("L1").Select
    Selection.AutoFill Destination:=Range("L1:L" & jrng)
    Range("L1:L" & jrng).Select
    Selection.Copy
    Range("F1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Columns("G:M").Delete
End Sub

Private Sub 插入1()
 'irow = Range("h65536").End(xlUp).Row
 Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Copy Range("h1")
    irow = Range("h65536").End(xlUp).Row
    Range("h1:h" & irow).Value = "1"
End Sub

Private Sub 替换()
    Cells.Replace What:="宝贝属性：", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Private Sub 清空()
    Columns("G:G").ClearContents
End Sub

Private Sub 复制()
    Sheets(1).Range("b2:l500").ClearContents
    irow = Range("h65536").End(xlUp).Row
    Sheets(Sheets.Count).Range("c1:h" & irow).Copy Sheets(1).Range("b2")
    
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub

Sub 无订单号打单()
    If Range("a1") = "打印状态" Then
        Call 删除
        Call 分列
        Call 合并
        Call 插入1
        Call 复制
    End If
End Sub

Sub 无订单号空包()
    If Range("a1") = "打印状态" Then
        Call 删除
        Call 分列
        Call 合并
        Call 插入1
        Call 清空
        Call 复制
    End If
End Sub



