Attribute VB_Name = "ģ��2"
Sub ����ϲ�()
'����Ի������
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'�½�һ��������
Dim newwb As Workbook
Set newwb = Workbooks.Add
With fd
If .Show = -1 Then
'���嵥���ļ�����
Dim vrtSelectedItem As Variant
'����ѭ������
Dim i As Integer
i = 1
'��ʼ�ļ�����
For Each vrtSelectedItem In .SelectedItems
'�򿪱��ϲ�������
Dim tempwb As Workbook
Set tempwb = Workbooks.Open(vrtSelectedItem)
'���ƹ�����
tempwb.Worksheets.Copy Before:=newwb.Worksheets(i)
'���¹������Ĺ��������ָĳɱ����ƹ������ļ��������Ӧ����xls�ļ�����Excel97-2003���ļ��������Excel2007����Ҫ�ĳ�xlsx
newwb.Worksheets(i).Name = VBA.Replace(tempwb.Name, ".xlsx", "")
'�رձ��ϲ�������
tempwb.Close SaveChanges:=False
i = i + 1
Next vrtSelectedItem
End If
End With
Set fd = Nothing
End Sub

