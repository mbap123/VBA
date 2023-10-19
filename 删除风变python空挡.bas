Attribute VB_Name = "É¾³ý·ç±äpython¿Õµ²"


Sub É¾³ý·ç±äpython¿Õµ²()
Dim arr, brr()
arr = Range("a1:a" & Cells(Rows.Count, 1).End(xlUp).Row)
arr = WorksheetFunction.Transpose(arr)

n = 1
For i = 1 To UBound(arr)
    If Application.WorksheetFunction.Count(arr(i)) = 0 And Len(arr(i)) > 1 Then
    n = n + 1
    ReDim Preserve brr(1 To n)
    brr(n) = arr(i)
    End If
Next

[h1:z100].Clear
[h1].Resize(UBound(brr), 1) = WorksheetFunction.Transpose(brr)

End Sub
