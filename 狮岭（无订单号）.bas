Attribute VB_Name = "ʨ��"

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
    Columns("C:C").Select
    Selection.Cut
    Columns("e:e").Select
    Selection.Insert Shift:=xlToRight
End Sub

Private Sub �滻()
    Cells.Replace What:="�������ԣ�", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Private Sub ����()
    Sheets(1).Range("a1:g500").ClearContents
    irow = Range("d65536").End(xlUp).Row
    Sheets(Sheets.Count).Range("c1:d" & irow).Copy Sheets(1).Range("a1")
    
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
End Sub


Sub ʨ��()
    If Range("a1") = "��ӡ״̬" Then
        Call ɾ��
        Call ����
        Call �滻
        Call ����
    End If
End Sub



