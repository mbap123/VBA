Attribute VB_Name = "模块2"
Sub 导入合并()
'定义对话框变量
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'新建一个工作簿
Dim newwb As Workbook
Set newwb = Workbooks.Add
With fd
If .Show = -1 Then
'定义单个文件变量
Dim vrtSelectedItem As Variant
'定义循环变量
Dim i As Integer
i = 1
'开始文件检索
For Each vrtSelectedItem In .SelectedItems
'打开被合并工作簿
Dim tempwb As Workbook
Set tempwb = Workbooks.Open(vrtSelectedItem)
'复制工作表
tempwb.Worksheets.Copy Before:=newwb.Worksheets(i)
'把新工作簿的工作表名字改成被复制工作簿文件名，这儿应用于xls文件，即Excel97-2003的文件，如果是Excel2007，需要改成xlsx
newwb.Worksheets(i).Name = VBA.Replace(tempwb.Name, ".xlsx", "")
'关闭被合并工作簿
tempwb.Close SaveChanges:=False
i = i + 1
Next vrtSelectedItem
End If
End With
Set fd = Nothing
End Sub

