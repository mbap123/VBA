Attribute VB_Name = "模块1"
Sub 删除非指定列()
    Dim ws As Worksheet
    Dim headerCell As Range
    Dim deleteColumns As Range
    Dim columnHeader As String
    
    ' 指定列标签
    Dim specifiedColumns As Variant
    specifiedColumns = Array("订单状态", "商品ID", "商家备注", "售后状态")
    
    ' 设置工作表（根据实际情况调整工作表名称）
    'Set ws = ThisWorkbook.Sheets("Sheet1")  ' 修改为你的工作表名称
    Set ws = ActiveSheet
    ' 遍历列标签，构建要删除的列的范围
    For Each headerCell In ws.Rows(1).Cells
        columnHeader = headerCell.value
        If IsError(Application.Match(columnHeader, specifiedColumns, 0)) Then
            If deleteColumns Is Nothing Then
                Set deleteColumns = headerCell
            Else
                Set deleteColumns = Union(deleteColumns, headerCell)
            End If
        End If
    Next headerCell
    
    ' 删除非指定列
    If Not deleteColumns Is Nothing Then
        deleteColumns.EntireColumn.Delete
    End If
End Sub



