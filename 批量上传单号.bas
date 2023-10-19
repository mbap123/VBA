Attribute VB_Name = "批量单号"

Private Sub 删除空格2()
    Range("g1").Delete
    maxh = Sheets(Sheets.Count).Range("d65536").End(3).Row
    Range("d1:d" & maxh).SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
    Range("M2") = "=MID(RC[-6],4,16)"
    Range("M2").AutoFill Range("m2:m" & Range("d65536").End(xlUp).Row)
    
    Range("l2:l" & Range("d65536").End(xlUp).Row) = "圆通快递"
    Sheets(Sheets.Count).Range("m2:m" & Range("d65536").End(xlUp).Row).Copy
    Sheets(Sheets.Count).Range("m2:m" & Range("d65536").End(xlUp).Row). _
    PasteSpecial Paste:=xlPasteValues
    
    Sheets(Sheets.Count).Range("d2:d" & Range("d65536").End(xlUp).Row).Copy
    Sheets(Sheets.Count).Range("k2:k" & Range("d65536").End(xlUp).Row). _
    PasteSpecial Paste:=xlPasteValues
    Range("C1").Select
    
 End Sub

Private Sub 删除空格()
    Range("A1:A2").Select
    Selection.Insert Shift:=xlToRight
    Range("g1").Delete
    
    maxh = Sheets(Sheets.Count).Range("d65536").End(3).Row
    Range("d1:d" & maxh).SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
    Range("M1") = "=MID(RC[-6],4,16)"
    Range("M1").AutoFill Range("m1:m" & Range("d65536").End(xlUp).Row)
    
    Range("l1:l" & Range("d65536").End(xlUp).Row) = "圆通快递"
    Sheets(Sheets.Count).Range("m1:m" & Range("d65536").End(xlUp).Row).Copy
    Sheets(Sheets.Count).Range("m1:m" & Range("d65536").End(xlUp).Row). _
    PasteSpecial Paste:=xlPasteValues
    
    Sheets(Sheets.Count).Range("d1:d" & Range("d65536").End(xlUp).Row).Copy
    Sheets(Sheets.Count).Range("k1:k" & Range("d65536").End(xlUp).Row). _
    PasteSpecial Paste:=xlPasteValues
    Range("C1").Select
    
 End Sub
Private Sub 复制()
    Sheets(1).Range("a2:l500").ClearContents
    Sheets(Sheets.Count).Range("k1:m" & Range("d65536").End(xlUp).Row).Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        '删除表二数据
        Sheets(2).Range("a1:n500").ClearContents
        Sheets(2).Columns("C:D").NumberFormatLocal = "@"
        Sheets("Sheet2").Activate
    Range("a1").Select
       
    
End Sub

Private Sub 筛选()
    'Sheets(1).Range("a2:l500").ClearContents
    'Sheets(Sheets.Count).Range("a1").EntireRow.Delete
    'Sheets(Sheets.Count).Range("k2:m" & Range("d65536").End(xlUp).Row).Delete
    Sheets(Sheets.Count).Range("N:Q").Delete
    Sheets(Sheets.Count).Range("F:L").Delete
    Sheets(Sheets.Count).Range("A:D").Delete
    
    
    Range("A2").Select
    
End Sub

Sub 批量上传()
    Call 删除空格
    Call 复制
    
End Sub


Sub 批量生成两列()
    Call 删除空格
    Call 筛选

End Sub
