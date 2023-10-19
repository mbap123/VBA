Attribute VB_Name = "筛选编号爱用交易"


Function vv(Rng1 As Variant, ze1 As String) '传入各种型
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = False
    .Pattern = ze1 '写正则表达式
  Set m1 = .Execute(Rng1)
  End With
  vv = m1(0) '为列表，即使只有一个值，也需要以数组格式复制，如果省略括号的0，则报错
End Function

Sub 筛选编号爱用交易()
arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
arr = WorksheetFunction.Transpose(arr)
arr1 = Filter(arr, "编号:", True)

For i = 0 To UBound(arr1)
        arr1(i) = vv(arr1(i), "\d{6,}")
Next

Sheets(1).UsedRange.ClearContents
Sheets(1).[a1].Resize(UBound(arr1), 1) = WorksheetFunction.Transpose(arr1)
Sheets(2).UsedRange.Clear
Sheets(1).Select

End Sub


