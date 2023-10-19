Attribute VB_Name = " ®¡Î"

Private Sub …æ≥˝()
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

Private Sub ≤Â»Î()
    Columns("C:C").Select
    Selection.Cut
    Columns("e:e").Select
    Selection.Insert Shift:=xlToRight
End Sub

Private Sub ÃÊªª()
    Cells.Replace What:="±¶±¥ Ù–‘£∫", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Private Sub ∏¥÷∆()
    Sheets(1).Range("a1:g500").ClearContents
    irow = Range("d65536").End(xlUp).Row
    Sheets(Sheets.Count).Range("c1:d" & irow).Copy Sheets(1).Range("a1")
    
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub


Sub  ®¡Î()
    If Range("a1") = "¥Ú”°◊¥Ã¨" Then
        Call …æ≥˝
        Call ≤Â»Î
        Call ÃÊªª
        Call ∏¥÷∆
    End If
End Sub



