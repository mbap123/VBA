Attribute VB_Name = "�������ſ��"
Private Sub ����()
    Dim df As Range, d$, fz As Range, zt As Range
    Set df = Range("a:a").Find("���������*", , xlValues, 2)
    d = df.Address
    Do
        Set df = Range("a:a").FindNext(df)
        Set fz = df.Resize(1, 4)
        Set zt = Cells(Rows.Count, "d").End(xlUp).Offset(1, 0)
        fz.Copy zt
    Loop Until df.Address = d
    
End Sub

Private Sub �ջ�()
    Dim sh As Range, s$, fzs As Range, zts As Range
    Set sh = Range("a:a").Find("�ջ���ַ*", , xlValues, 2)
    s = sh.Address
    Do
        Set sh = Range("a:a").FindNext(sh)
        Set fzs = sh.Resize(1, 4)
        Set zts = Cells(Rows.Count, "e").End(xlUp).Offset(1, 0)
        fzs.Copy zts
    Loop Until sh.Address = s
    
End Sub

Private Sub ����()
    Dim bb As Range, b$, fzb As Range, ztb As Range
    Set bb = Range("a:a").Find("��������*", , xlValues, 2)
    b = bb.Address
    Do
        Set bb = Range("a:a").FindNext(bb)
        Set fzb = bb.Resize(1, 4)
        Set ztb = Cells(Rows.Count, "k").End(xlUp).Offset(1, 0)
        fzb.Copy ztb
    Loop Until bb.Address = b
    
End Sub

Private Sub ɸѡ()
    Range("F2") = "=MID(RC[-2],7,19)"
    Range("f2").AutoFill Range("f2:f" & Range("d65536").End(xlUp).Row)
    
    Range("G2") = "=MID(RC[-2],FIND(""��"",RC[-2])+1,FIND("","",RC[-2])-FIND(""��"",RC[-2])-1)"
    Range("g2").AutoFill Range("g2:g" & Range("d65536").End(xlUp).Row)
    
    Range("i2") = "=TRIM(MID(SUBSTITUTE(RC[-4],"","",REPT("" "",LEN(RC[-4]))),LEN(RC[-4]),LEN(RC[-4])))"
    Range("i2").AutoFill Range("i2:i" & Range("d65536").End(xlUp).Row)
    
    Range("j2") = "=TRIM(MID(SUBSTITUTE(RC[-5],"","",REPT("" "",LEN(RC[-5]))),2*LEN(RC[-5]),LEN(RC[-5])))"
    Range("j2").AutoFill Range("j2:j" & Range("d65536").End(xlUp).Row)
    
    Range("l2:l" & Range("d65536").End(xlUp).Row) = "1"
    
End Sub
Private Sub ����()
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
Private Sub ����2()
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

Sub �������Ŵ�()
    Call ����
    Call �ջ�
    Call ����
    Call ɸѡ
    Call ����
End Sub

Sub �������ſհ�()
    Call ����
    Call �ջ�
    Call ����
    Call ɸѡ
    Call ����2
End Sub
