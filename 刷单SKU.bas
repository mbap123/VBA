Attribute VB_Name = "刷单SKU"

Function zz(Rng1 As Range, ze1 As String) '传入单元格
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = False
    .Pattern = ze1 '写正则表达式
  Set m1 = .Execute(Rng1)
  End With
  zz = m1(0) '为列表，即使只有一个值，也需要以数组格式复制，如果省略括号的0，则报错
End Function

Function vv(Rng1 As Variant, ze1 As String) '传入各种型
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = False
    .Pattern = ze1 '写正则表达式
  Set m1 = .Execute(Rng1)
  End With
  vv = m1(0) '为列表，即使只有一个值，也需要以数组格式复制，如果省略括号的0，则报错
End Function

Function vva(Rng1 As Variant, ze1 As String, a$) '传入各种型，匹配正则第几个
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = True
    .Pattern = ze1 '写正则表达式
  Set m1 = .Execute(Rng1)
  End With
  vva = m1(a) '为列表，即使只有一个值，也需要以数组格式复制，如果省略括号的0，则报错
End Function

Function vvd(Rng1 As Range, ze1 As String, i$) '传入单元格，匹配正则第几个
    Set regx1 = CreateObject("vbscript.regexp")
  With regx1
    .Global = True
    .Pattern = ze1 '写正则表达式
  Set m1 = .Execute(Rng1)
  End With
  vvd = m1(i) ' & m1(1) 为列表，即使只有一个值，也需要以数组格式复制，如果省略括号的0，则报错
End Function

Private Sub 查询是否有合单()
Dim rng As Range, rngs As String
On Error Resume Next
Set rng = ActiveSheet.UsedRange.Find("取消合单")
rngs = rng.Address(0, 0)
Do
Set rng = ActiveSheet.UsedRange.FindNext(rng)
    If zz(rng.Offset(1, 0), "[一-龢]{3}") <> rng.Offset(-2, 0) Then
        rng.Offset(1, 0) = rng.Offset(-2, 0) & rng.Offset(1, 0)
    End If
Loop Until rng.Address(0, 0) = rngs
End Sub

Private Sub 筛选初数据()
On Error Resume Next

Dim arr(), brr(), m, n As Integer
arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
arr = WorksheetFunction.Transpose(arr)
n = 0
    For i = 1 To UBound(arr)
        If arr(i) Like "???编号:*" Then
            ReDim Preserve brr(UBound(arr), 8)
            brr(n, 0) = vv(arr(i), "\d{4}-\d{2}-\d{2}")
                If Len(arr(i + 2)) > 8 Then
                    brr(n, 1) = arr(i + 3)
                Else
                    brr(n, 1) = arr(i + 2)
                End If
            brr(n, 2) = vv(arr(i), "[一-龢]{3}编号")
            brr(n, 3) = vv(arr(i), "\d{5,}")
        ElseIf arr(i) Like "宝贝属性：*" Then
                brr(n, 4) = vv(arr(i), "[^宝贝属性：].+")
        ElseIf arr(i) Like "买家档案" Then
                brr(n, 5) = arr(i - 1)
        ElseIf arr(i) Like "收货地址：" Then
                brr(n, 6) = vva(arr(i + 1), "[^,]+", 0)
                brr(n, 7) = vva(arr(i + 3), "[一-龢]{2,} ", 0)
                n = n + 1
        End If
           
    Next
Columns("g:o").Clear: Columns("j:j").NumberFormatLocal = "@": Columns("g:g").NumberFormatLocal = "mm-dd;@"
Range(Cells(1, "g"), Cells(UBound(brr) + 1, "o")) = brr
Columns("g:p").AutoFit
Erase arr: Erase brr
End Sub

Private Sub 复制到新表()
ActiveSheet.Copy after:=ActiveSheet
Columns("A:f").Delete
Cells.EntireColumn.AutoFit
'ActiveSheet.Name = "筛选"
Rows(1).Insert
[a1:j1] = Array("时间", "简称", "状态", "编号", "退货", "旺旺", "姓名", "省份", "金额", "成本")
Rows(1).Interior.Color = 14277081
Rows(1).Font.Size = 13
Rows(1).Font.Bold = True
Columns(1).HorizontalAlignment = xlRight: Columns(3).HorizontalAlignment = xlCenter
Columns(2).HorizontalAlignment = xlRight
Rows(1).VerticalAlignment = xlCenter: Rows(1).HorizontalAlignment = xlCenter
End Sub

Private Sub 整理工作表()
Dim rng As Range
arr = Range("c1:f" & Range("c65536").End(xlUp).Row)
For i = 1 To UBound(arr)
    If arr(i, 1) Like "*编号" Then
        arr(i, 1) = Replace(arr(i, 1), "编号", "")
    End If
Next

Range("c1").Resize(UBound(arr), 4).Value = arr
Cells.EntireColumn.AutoFit
Columns("F:G").ColumnWidth = 9
Columns(1).ColumnWidth = 11: Columns(2).ColumnWidth = 12: Columns(3).ColumnWidth = 13
End Sub

Private Sub 统计重复累加()
Dim arr, addrr, a00rr, brr, ir, lie, i, yyc As Integer
Columns("k:z").ClearContents
Set dd = CreateObject("scripting.dictionary")
Set d = CreateObject("scripting.dictionary")
Set y = CreateObject("scripting.dictionary")
arr = [a1].CurrentRegion
    For j = 2 To UBound(arr)
        dd(arr(j, 3)) = dd(arr(j, 3)) + 1
        d(arr(j, 2)) = d(arr(j, 2)) + 1
    Next
        [j2].Resize(dd.Count) = WorksheetFunction.Transpose(dd.keys)
        [k1].Resize(1, d.Count) = WorksheetFunction.Transpose(WorksheetFunction.Transpose(d.keys))
        brr = Range("k1", Cells(1, Columns.Count).End(xlToLeft))
        brr = WorksheetFunction.Transpose(WorksheetFunction.Transpose(brr))
            If VarType(brr) = vbString Then '如果是字符串的话转化为数组
            brr = Split(brr, "")
            End If
    ir = 1
    For Each mulu In brr
        For j = 2 To UBound(arr)
            If arr(j, 2) = mulu Then
                y(arr(j, 3)) = y(arr(j, 3)) + 1
            End If
        Next
        For Each rng In Range("j2：j" & Range("j65536").End(xlUp).Row)
            For Each i In y.keys()
                If rng = i Then
                rng.Offset(0, ir) = y(i)
                End If
            Next
        Next
        y.RemoveAll
        ir = ir + 1
    Next
    Columns("J").Insert: Columns("J").ColumnWidth = 4
'    Columns("K").Font.Bold = True
For Each rng In Range("k2:k" & Range("k65536").End(xlUp).Row)
    If rng = "待付款" Or rng = "未付款" Then
    rng.Font.Color = 255
    End If
Next

'
'二次统计表格*********************************************************************************************
Range(Range("k2").End(xlDown).Offset(1, 0), Range("z100")).ClearContents
'提取说明
With Sheets("00")
    a00rr = .Range(.Range("D1"), .Range("E" & Rows.Count).End(xlUp))
End With
    yy = Cells(1, Columns.Count).End(xlToLeft).Column
    xx = Cells(Rows.Count, "k").End(xlUp).Row
    addrr = Range("K1", Cells(xx, yy))
    addrr = WorksheetFunction.Transpose(addrr)
    
    x = Cells(Rows.Count, "k").End(xlUp).Offset(2, 0).Address(0, 0)
    '写说明
    Range(x).Resize(UBound(a00rr), 2) = a00rr
    '字体格式
    Range(x).Offset(0, 1).Resize(UBound(a00rr), 1).Font.Color = 10921638
    
    bt = Cells(Rows.Count, "k").End(xlUp).Offset(2, 1).Address(0, 0)
    Range("L1", Cells(1, Columns.Count).End(xlToLeft)).Copy Range(bt)
    y = Cells(Rows.Count, "k").End(xlUp).Offset(3, 0).Address(0, 0)
    Range(Range(y).Offset(-1, 1), Cells(Range(y).Offset(-1, 1).Row, _
                                Columns.Count).End(xlToLeft)).Interior.Color = 11389944

    Range(y).Resize(4, 1) = WorksheetFunction.Transpose(Array("退款", "成功", "待定", "总计"))
    Range(y).Resize(4, 1).Font.Bold = True
    
    yyc = Cells(1, Columns.Count).End(xlToLeft).Column
    lie = 2
For Each rng In Range(Range(y).Offset(-1, 0), Cells(Range(y).Row - 1, yyc))
    If rng <> "" Then
    For i = 1 To UBound(addrr, 2)

                If addrr(1, i) = "退款完" Then
                    rng.Offset(1, 0) = addrr(lie, i)
                ElseIf addrr(1, i) = "退款中" Then
                    rng.Offset(1, 0) = rng.Offset(1, 0) + addrr(lie, i)
                ElseIf addrr(1, i) = "已成功" Then
                    rng.Offset(2, 0) = addrr(lie, i)
                ElseIf addrr(1, i) = "待评价" Then
                    rng.Offset(2, 0) = rng.Offset(2, 0) + addrr(lie, i)
                ElseIf addrr(1, i) = "待发货" Then
                    rng.Offset(3, 0) = addrr(lie, i)
                ElseIf addrr(1, i) = "已发货" Then
                    rng.Offset(3, 0) = rng.Offset(3, 0) + addrr(lie, i)
                End If

                If addrr(1, i) <> "待付款" And i <> 1 And addrr(1, i) <> "未付款" Then
                    rng.Offset(4, 0) = rng.Offset(4, 0) + addrr(lie, i)
                End If
    Next
    lie = lie + 1
    End If
Next

    
    
End Sub
Private Sub 匹配成本()
Dim rng As Range, j As Range, i, lRow As Long, Dict As Object, myArray As Variant
    With Sheets("00")
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        myArray = .Range("A1").Resize(lRow, 2)
        
        Set Dict = CreateObject("scripting.dictionary")
        Dict.CompareMode = vbTextCompare

        For i = 2 To UBound(myArray)
        Dict(myArray(i, 1)) = i
        Next
    End With
    With ActiveSheet
        Set rng = .Range(.Range("B2"), .Range("B" & Rows.Count).End(xlUp))
        For Each j In rng
            If Dict.exists(j.Value) Then
                j.Offset(0, 7) = myArray(Dict(j.Value), 2)
            End If
        Next j
    End With
End Sub
Sub 刷单SKU()
Call 查询是否有合单

Call 筛选初数据
Call 复制到新表
Call 整理工作表
Call 统计重复累加
'Call 匹配成本
End Sub








