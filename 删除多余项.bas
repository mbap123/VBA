Attribute VB_Name = "删除多余项"
Private Sub 生成字符串()
    
    Dim i%
    For i = 1 To Cells(Rows.Count, "A").End(xlUp).Row
        Cells(i, 10) = Len(Cells(i, 1))
Next
End Sub

Private Sub 删除多余()
    Dim n%
    For n = Cells(Rows.Count, "J").End(xlUp).Row To 1 Step -1
    If Range("j" & n) < 40 Then
    Range("j" & n).EntireRow.Delete
    End If
    Next
End Sub

Sub 一键除余项()
    Dim n%
    For n = Cells(Rows.Count, "A").End(xlUp).Row To 1 Step -1
        If Len(Cells(n, 1)) < 40 Then
        Range("j" & n).EntireRow.Delete
        End If
    Next
End Sub
