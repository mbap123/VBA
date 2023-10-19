Attribute VB_Name = "微信聊天记录筛选旺旺号"
Sub 微信筛选()
Dim brr(), crr()
arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
arr = WorksheetFunction.Transpose(arr)
n = 0
    For i = 1 To UBound(arr)
    '    If Len(arr(i)) > 5 Then
        If arr(i) Like "*，*" Then
        n = n + 1
        ReDim Preserve brr(1 To n)
        brr(n) = arr(i)
        End If
    Next
Columns("C:U").Clear
Range("c1:c" & UBound(brr)) = WorksheetFunction.Transpose(brr)
brr = WorksheetFunction.Transpose(WorksheetFunction.Transpose(brr))
n = 0
For i = 1 To UBound(brr)
    n = n + 1
        ReDim Preserve crr(1 To n)
        crr(n) = Split(brr(i), "，")(0)
'        crr(n) = Split(brr(i), "，")(1)

Next

Range("h1:h" & UBound(brr)) = WorksheetFunction.Transpose(crr)
End Sub

