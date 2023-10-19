Attribute VB_Name = "°¢Àï°Í°ÍÒþ²ØÁÐ"
Sub Òþ²ØÁÐ()

ActiveSheet.Copy after:=ActiveSheet
Columns("S:AD").Delete
Columns("P:Q").Delete
Columns("F:M").Delete
Columns("A:D").Delete

    
    Columns("C:C").Cut Columns("G:G")
    Columns("D:D").Cut Columns("H:H")
    Columns("C:D").Delete
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 17
    Columns("D:D").ColumnWidth = 23
    Columns("E:F").Font.Color = 10921638
End Sub
