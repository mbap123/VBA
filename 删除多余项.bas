Attribute VB_Name = "ɾ��������"
Private Sub �����ַ���()
    
    Dim i%
    For i = 1 To Cells(Rows.Count, "A").End(xlUp).Row
        Cells(i, 10) = Len(Cells(i, 1))
Next
End Sub

Private Sub ɾ������()
    Dim n%
    For n = Cells(Rows.Count, "J").End(xlUp).Row To 1 Step -1
    If Range("j" & n) < 40 Then
    Range("j" & n).EntireRow.Delete
    End If
    Next
End Sub

Sub һ��������()
    Dim n%
    For n = Cells(Rows.Count, "A").End(xlUp).Row To 1 Step -1
        If Len(Cells(n, 1)) < 40 Then
        Range("j" & n).EntireRow.Delete
        End If
    Next
End Sub
