Attribute VB_Name = "������������"
Function vv(Rng1 As Variant, ze1 As String) '���������
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = False
    .Pattern = ze1 'д������ʽ
  Set m1 = .Execute(Rng1)
  End With
  vv = m1(0) 'Ϊ�б���ʹֻ��һ��ֵ��Ҳ��Ҫ�������ʽ���ƣ����ʡ�����ŵ�0���򱨴�
End Function
Function �ۼ�()
Dim sum As Integer, n As Integer, b2rr(), sumi As Integer
sum = [G3]: n = 10
x = 1
For i = 1 To 10
    ReDim Preserve b2rr(1 To 10)
    b2rr(i) = i
    sumi = sumi + b2rr(i)
    x = x + 1
        If sumi >= sum Then GoTo tt
        
Next
tt:
i = 1
If sumi < sum Then
        Do Until sumi = sum
                b2rr(i) = b2rr(i) + 1
                i = i + 1
                sumi = sumi + 1
            If i = 11 Then
                i = 1
            End If
        Loop
End If
If sumi > sum Then
    i = UBound(b2rr)
    Do Until sumi = sum
        b2rr(i) = b2rr(i) - 1
        sumi = sumi - 1
        i = i - 1
        If i = 1 Then
                i = UBound(b2rr)
            End If
    Loop
End If

�ۼ� = b2rr
End Function
Function ƽ��()

Dim sum As Integer, n As Integer, pin(), sumi As Integer
sum = [G3]: n = 10
avg = Int(sum / n)
For i = 1 To n
    ReDim Preserve pin(1 To n)
    pin(i) = avg
    sumi = sumi + pin(i)
Next

If sumi < sum Then
    i = 1
    Do Until sumi = sum
        pin(i) = pin(i) + 1
        sumi = sumi + 1
        i = i + 1
    Loop
End If

ƽ�� = pin
End Function

Sub ����()
Dim rng As Range
ActiveSheet.Name = 0

For Each rng In Range("r1:R" & Range("r65536").End(xlUp).Row)
    If rng <> "" Then
        x = rng
        ActiveSheet.Copy before:=ActiveSheet
        ActiveSheet.Name = rng.Offset(0, -2)
        ActiveSheet.Range("e1") = rng.Offset(0, -2)
    End If
Next
End Sub

Sub ����2��()

On Error Resume Next
Dim arr, rng As Range, arr1, brr(), crr(), c1rr(), d1rr()

Set d = CreateObject("scripting.dictionary")
arr = Sheets("�ʱ�").UsedRange
'arr = Sheets("�ʱ�").Range("a1:b" & Sheets("�ʱ�").Range("b65536").End(xlUp).Row)
For j = 1 To UBound(arr)
    d(arr(j, 1)) = arr(j, 2)
Next

For Each rng In Range("r1:R" & Range("r65536").End(xlUp).Row)
    If rng <> "" Then
        ActiveSheet.Copy before:=ActiveSheet
        ActiveSheet.Name = rng.Offset(0, -2)
        ActiveSheet.Range("e1") = rng.Offset(0, -2)
        x = rng.Offset(0, -2)
        y = vv(x, "[^\d]+")
        arr1 = Filter(d.keys, y, True)
            For i = 0 To UBound(arr1)
                ReDim Preserve brr(1 To 10, 1 To 5)
                brr(i + 1, 5) = arr1(i)
                brr(i + 1, 1) = d(brr(i + 1, 5))
                brr(i + 1, 2) = ƽ��(i + 1)
'                brr(i + 1, 2) = �ۼ�(i + 1)
'                brr(i + 1, 3) = Round(brr(i + 1, 2) / [G3] * [H3])
'                brr(i + 1, 4) = Round(brr(i + 1, 2) / [G3] * [i3])
'                =ROUND((B1/$G$3)*$H$3,0),=ROUND((B1/$G$3)*$I$3,0)
                brr(i + 1, 3) = "=ROUND((B" & i + 1 & "/$G$3)*$H$3,0)"
                brr(i + 1, 4) = "=ROUND((B" & i + 1 & "/$G$3)*$I$3,0)"
            Next
            Columns("A:E").ClearContents
            Range("a65536").End(xlUp).Offset(0, 0).Resize(UBound(brr), 5) = brr
            Erase brr
    End If
Next
End Sub

Sub ����()
Range("b1:b" & Range("b65536").End(xlUp).Row).Sort Range("B1"), xlDescending 'xlAscending��
Range("b1:b" & Range("b65536").End(xlUp).Row).Select
End Sub

Sub ɾ��������()
Dim sht As Worksheet
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        With sht
            If .Name = "0" Or .Name = "�ʱ�" Or .Name = "�ʱ���" Or .Name = "���ڵ�" Then
                Exit Sub
            Else
                sht.Delete
            End If
        End With
    Next
    Application.DisplayAlerts = True
End Sub

Sub ΢��ɸѡ����()
Dim brr(), crr()
arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
arr = WorksheetFunction.Transpose(arr)
n = 0
    For i = 1 To UBound(arr)
    '    If Len(arr(i)) > 5 Then
        If arr(i) Like "*��*" Then
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
        crr(n) = Split(brr(i), "��")(0)
'        crr(n) = Split(brr(i), "��")(1)

Next

Range("h1:h" & UBound(brr)) = WorksheetFunction.Transpose(crr)
End Sub


