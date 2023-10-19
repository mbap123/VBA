Attribute VB_Name = "带订单号快递"
Private Sub 待发()
    Dim df As Range, d$, fz As Range, zt As Range
    Set df = Range("a:a").Find("待发货编号*", , xlValues, 2)
    d = df.Address
    Do
        Set df = Range("a:a").FindNext(df)
        Set fz = df.Resize(1, 4)
        Set zt = Cells(Rows.Count, "d").End(xlUp).Offset(1, 0)
        fz.Copy zt
    Loop Until df.Address = d
    
End Sub

Private Sub 收货()
    Dim sh As Range, s$, fzs As Range, zts As Range
    Set sh = Range("a:a").Find("收货地址*", , xlValues, 2)
    s = sh.Address
    Do
        Set sh = Range("a:a").FindNext(sh)
        Set fzs = sh.Resize(1, 4)
        Set zts = Cells(Rows.Count, "e").End(xlUp).Offset(1, 0)
        fzs.Copy zts
    Loop Until sh.Address = s
    
End Sub

Private Sub 宝贝()
    Dim bb As Range, b$, fzb As Range, ztb As Range
    Set bb = Range("a:a").Find("宝贝属性*", , xlValues, 2)
    b = bb.Address
    Do
        Set bb = Range("a:a").FindNext(bb)
        Set fzb = bb.Resize(1, 4)
        Set ztb = Cells(Rows.Count, "k").End(xlUp).Offset(1, 0)
        fzb.Copy ztb
    Loop Until bb.Address = b
    
End Sub

Private Sub 筛选()
    Range("F2") = "=MID(RC[-2],7,19)"
    Range("f2").AutoFill Range("f2:f" & Range("d65536").End(xlUp).Row)
    
    Range("G2") = "=MID(RC[-2],FIND(""："",RC[-2])+1,FIND("","",RC[-2])-FIND(""："",RC[-2])-1)"
    Range("g2").AutoFill Range("g2:g" & Range("d65536").End(xlUp).Row)
    
    Range("i2") = "=TRIM(MID(SUBSTITUTE(RC[-4],"","",REPT("" "",LEN(RC[-4]))),LEN(RC[-4]),LEN(RC[-4])))"
    Range("i2").AutoFill Range("i2:i" & Range("d65536").End(xlUp).Row)
    
    Range("j2") = "=TRIM(MID(SUBSTITUTE(RC[-5],"","",REPT("" "",LEN(RC[-5]))),2*LEN(RC[-5]),LEN(RC[-5])))"
    Range("j2").AutoFill Range("j2:j" & Range("d65536").End(xlUp).Row)
    
    Range("l2:l" & Range("d65536").End(xlUp).Row) = "1"
    
End Sub
Private Sub 复制()
    Sheets(1).Range("a2:l500").ClearContents
    Sheets(Sheets.Count).Range("f2:l" & Range("l65536").End(xlUp).Row).Copy
    Sheets("Sheet1").Activate
    Range("a2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
    Range("c2:c" & Range("a65536").End(xlUp).Row).ClearContents
    Range("a2").Select
    
End Sub
Private Sub 复制2()
    Sheets(1).Range("a2:l500").ClearContents
    Sheets(Sheets.Count).Range("f2:l" & Range("l65536").End(xlUp).Row).Copy
    Sheets("Sheet1").Activate
    Range("a2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
    Application.DisplayAlerts = True
    Range("c2:c" & Range("a65536").End(xlUp).Row).ClearContents
    Range("f2:f" & Range("a65536").End(xlUp).Row).ClearContents
    Range("a2").Select
    
End Sub

Sub 带订单号打单()
    Call 待发
    Call 收货
    Call 宝贝
    Call 筛选
    Call 复制
End Sub

Sub 带订单号空包()
    Call 待发
    Call 收货
    Call 宝贝
    Call 筛选
    Call 复制2
End Sub
