Attribute VB_Name = "���ļ���"

Sub �½��ļ���()
On Error Resume Next

Dim path As String
path = "C:\Users\32897\Desktop\the row ns park"
For Each r In Array("��ͼ", "��ɫ", "Model", "ϸ��", "��ѡ")
    MkDir (path & Application.PathSeparator & r)
Next
End Sub


Sub �½��ļ���2()
On Error Resume Next

Dim path As String, ne As String
path = "C:\Users\32897\Desktop\���ۻ���"

arr = Range("a1:d" & Range("d65536").End(xlUp).Row)
For i = 1 To UBound(arr)
    ne = arr(i, 1) & "_" & arr(i, 2) & "_" & arr(i, 3) & "_" & arr(i, 4)
    MkDir (path & Application.PathSeparator & ne)
Next

End Sub
