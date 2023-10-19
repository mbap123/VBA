Attribute VB_Name = "模块4"
Sub CreateMenu()
Sheets.Add(Before:=Sheets(1)).Name = "目录"  '新建一个目录工作表
Worksheets("目录").Activate
    For i = 1 To Sheets.Count
        Cells(i, 1) = Sheets(i).Name  '将其他工作表名称分别填入单元格中
        If i <> 1 Then
            Cells(i, 1).Select
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & Sheets(i).Name & "'!A1", TextToDisplay:=Cells(i, 1).Value
                '创建超链接
        End If
    Next i
End Sub
