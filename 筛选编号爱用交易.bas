Attribute VB_Name = "ɸѡ��Ű��ý���"


Function vv(Rng1 As Variant, ze1 As String) '���������
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = False
    .Pattern = ze1 'д������ʽ
  Set m1 = .Execute(Rng1)
  End With
  vv = m1(0) 'Ϊ�б���ʹֻ��һ��ֵ��Ҳ��Ҫ�������ʽ���ƣ����ʡ�����ŵ�0���򱨴�
End Function

Sub ɸѡ��Ű��ý���()
arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
arr = WorksheetFunction.Transpose(arr)
arr1 = Filter(arr, "���:", True)

For i = 0 To UBound(arr1)
        arr1(i) = vv(arr1(i), "\d{6,}")
Next

Sheets(1).UsedRange.ClearContents
Sheets(1).[a1].Resize(UBound(arr1), 1) = WorksheetFunction.Transpose(arr1)
Sheets(2).UsedRange.Clear
Sheets(1).Select

End Sub


