Attribute VB_Name = "Ê¨Áë"

Private Sub É¾³ý()
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

Private Sub ²åÈë()
    Columns("C:C").Select
    Selection.Cut
    Columns("e:e").Select
    Selection.Insert Shift:=xlToRight
End Sub

Private Sub Ìæ»»()
    Cells.Replace What:="±¦±´ÊôÐÔ£º", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Private Sub Ìæ»»2()
    Range("A1").Select
    Cells.Replace What:="-ÈáÈíÕæÆ¤¡¾Ë«¼ç´ø¡¿", Replacement:="¡¾Ð¡Ë®Í°¡¿", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("B1").Select
End Sub

Private Sub ¸´ÖÆ()
    Sheets(1).Range("a1:g500").ClearContents
    irow = Range("d65536").End(xlUp).Row
    Sheets(Sheets.Count).Range("c1:d" & irow).Copy Sheets(1).Range("a1")
    
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub


Sub Ê¨Áë()
    If Range("a1") = "´òÓ¡×´Ì¬" Then
        Call É¾³ý
        Call ²åÈë
        Call Ìæ»»
        Call ¸´ÖÆ
    End If
End Sub



