Attribute VB_Name = "����"
Private Sub ɾ��()
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
Private Sub ����()
Application.DisplayAlerts = False
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
       :=Array(Array(2, 1)), TrailingMinusNumbers:=True
        Application.DisplayAlerts = True
End Sub

Private Sub �ϲ�()
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

Private Sub ����2()
  Columns("f:i").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Cut
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight
    Columns("p:p").Copy Range("s1")
   irow = Range("s65536").End(xlUp).Row
    Range("s1:s" & irow).Value = "1"
End Sub

Private Sub �滻()
    Cells.Replace What:="�������ԣ�", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Private Sub ���()
    Columns("G:G").ClearContents
End Sub

Private Sub ����()
    Sheets(1).Range("b3:v500").ClearContents
    irow = Range("s65536").End(xlUp).Row
    Sheets(Sheets.Count).Range("c1:s" & irow).Copy Sheets(1).Range("c3")
    
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub


Sub ����()
    If Range("a1") = "��ӡ״̬" Then
        Call ɾ��
        Call ����
        Call �ϲ�
        Call ����2
        Call ����
    End If
End Sub



