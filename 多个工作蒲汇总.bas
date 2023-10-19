Attribute VB_Name = "模块1"
Sub 多个Sheet到汇总()
    Dim i As Long, j As Long 'i是数据源表的最后一行，j是目标表(数据表)的最后一行
    Dim sht As Worksheet
    Dim lastRow As Long
    
    Application.ScreenUpdating = False '关闭屏幕刷新
    
    '先要删除所有数据
    Sheets("汇总").Range("A1:Z" & Rows.Count).ClearContents
    
    '复制数据
    For Each sht In Sheets
        If sht.Name <> "汇总" Then
            i = sht.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            lastRow = GetLastNonEmptyRow(Sheets("汇总"), 1, i)
            If lastRow = 1 And IsEmpty(Sheets("汇总").Range("A1")) Then
                lastRow = 0
            End If
            j = GetLastNonEmptyRow(Sheets("汇总"), 1, Columns.Count)
            sht.Range("A1:IV" & i).Copy Destination:=Sheets("汇总").Range("A" & j + 1)
        End If
    Next
    
    Application.ScreenUpdating = True '打开屏幕刷新
    
    MsgBox "执行完毕！"
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
 

