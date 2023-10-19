Attribute VB_Name = "shishi"
Sub test()
    Dim ss As Range, d$, fz As Range, zt As Range
    Set ss = Range("a:a").Find("±¶±¥ Ù–‘*", , xlValues, 2)
        d = ss.Address
        Do
        Set ss = Range("a:a").FindNext(ss)
        Set fz = ss.Resize(1, 4)
        Set zt = Cells(Rows.Count, "f").End(xlUp).Offset(1, 0)
        fz.Copy zt
        Loop Until ss.Address = d
End Sub

Sub shishi()
    Dim ss As Range, d$, fz As Range, zt As Range
    Set ss = Range("a:a").Find("±¶±¥ Ù–‘*", , xlValues, 2)
    d = ss.Address
        Do
            Set ss = Range("a:a").FindNext(ss)
            Set fz = ss.Resize(1, 4)
            Set zt = Cells(Rows.Count, "f").End(xlUp).Offset(1, 0)
            fz.Copy zt
        Loop Until ss.Address = d
End Sub


Sub shi2()
    Dim ss, df, sh As Range, d$, f$, s$, fz, fzf, fzs As Range, zt, ztf, zts As Range
    Set ss = Range("a:a").Find("±¶±¥ Ù–‘*", , xlValues, 2)
    Set df = Range("a:a").Find("¥˝∑¢ªı±‡∫≈*", , xlValues, 2)
    Set sh = Range("a:a").Find(" ’ªıµÿ÷∑*", , xlValues, 2)
    d = ss.Address
    f = df.Address
    s = sh.Address
        Do
        Do
        Do
            Set ss = Range("a:a").FindNext(ss)
            Set df = Range("a:a").FindNext(df)
            Set sh = Range("a:a").FindNext(sh)
            Set fz = ss.Resize(1, 1)
            Set fzf = df.Resize(1, 1)
            Set fzs = sh.Resize(1, 1)
            Set zt = Cells(Rows.Count, "d").End(xlUp).Offset(1, 0)
            Set ztf = Cells(Rows.Count, "e").End(xlUp).Offset(1, 0)
            Set zts = Cells(Rows.Count, "g").End(xlUp).Offset(1, 0)
            fz.Copy zt
            fzf.Copy ztf
            fzs.Copy zts
        Loop Until ss.Address = d
        Loop Until df.Address = f
        Loop Until sh.Address = s
        
End Sub

