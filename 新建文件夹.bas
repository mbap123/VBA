Attribute VB_Name = "新文件夹"

Sub 新建文件夹()
On Error Resume Next

Dim path As String
path = "C:\Users\32897\Desktop\the row ns park"
For Each r In Array("主图", "颜色", "Model", "细节", "备选")
    MkDir (path & Application.PathSeparator & r)
Next
End Sub


Sub 新建文件夹2()
On Error Resume Next

Dim path As String, ne As String
path = "C:\Users\32897\Desktop\评价汇总"

arr = Range("a1:d" & Range("d65536").End(xlUp).Row)
For i = 1 To UBound(arr)
    ne = arr(i, 1) & "_" & arr(i, 2) & "_" & arr(i, 3) & "_" & arr(i, 4)
    MkDir (path & Application.PathSeparator & ne)
Next

End Sub
