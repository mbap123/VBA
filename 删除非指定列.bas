Attribute VB_Name = "ģ��1"
Sub ɾ����ָ����()
    Dim ws As Worksheet
    Dim headerCell As Range
    Dim deleteColumns As Range
    Dim columnHeader As String
    
    ' ָ���б�ǩ
    Dim specifiedColumns As Variant
    specifiedColumns = Array("����״̬", "��ƷID", "�̼ұ�ע", "�ۺ�״̬")
    
    ' ���ù���������ʵ������������������ƣ�
    'Set ws = ThisWorkbook.Sheets("Sheet1")  ' �޸�Ϊ��Ĺ���������
    Set ws = ActiveSheet
    ' �����б�ǩ������Ҫɾ�����еķ�Χ
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
    
    ' ɾ����ָ����
    If Not deleteColumns Is Nothing Then
        deleteColumns.EntireColumn.Delete
    End If
End Sub



