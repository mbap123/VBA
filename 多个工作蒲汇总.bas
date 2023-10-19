Attribute VB_Name = "ģ��1"
Sub ���Sheet������()
    Dim i As Long, j As Long 'i������Դ������һ�У�j��Ŀ���(���ݱ�)�����һ��
    Dim sht As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False '�ر���Ļˢ��
    
    '��Ҫɾ����������
    Sheets("����").Range("A1:Z" & Rows.Count).ClearContents
    
    '��������
    For Each sht In Sheets
        If sht.Name <> "����" Then
            i = sht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            lastRow = GetLastNonEmptyRow(Sheets("����"), 1, i)
            If lastRow = 1 And IsEmpty(Sheets("����").Range("A1")) Then
                lastRow = 0
            End If
            j = GetLastNonEmptyRow(Sheets("����"), 1, Columns.Count)
            sht.Range("A1:IV" & i).Copy Destination:=Sheets("����").Range("A" & j + 1)
        End If
    Next
    
    Application.ScreenUpdating = True '����Ļˢ��
    
    MsgBox "ִ����ϣ�"
End Sub
 
Function GetLastNonEmptyRow(ws As Worksheet, startColumn As Long, endColumn As Long) As Long
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = ws.Cells(Rows.Count, startColumn).End(xlUp).Row
    For i = startColumn To endColumn
        lastRow = WorksheetFunction.Max(lastRow, ws.Cells(Rows.Count, i).End(xlUp).Row)
    Next i
    
    GetLastNonEmptyRow = lastRow
End Function
 

