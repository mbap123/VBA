Attribute VB_Name = "ģ��4"
Sub CreateMenu()
Sheets.Add(Before:=Sheets(1)).Name = "Ŀ¼"  '�½�һ��Ŀ¼������
Worksheets("Ŀ¼").Activate
    For i = 1 To Sheets.Count
        Cells(i, 1) = Sheets(i).Name  '���������������Ʒֱ����뵥Ԫ����
        If i <> 1 Then
            Cells(i, 1).Select
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & Sheets(i).Name & "'!A1", TextToDisplay:=Cells(i, 1).Value
                '����������
        End If
    Next i
End Sub
