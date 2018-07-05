Attribute VB_Name = "GetFolderHierarchy"
Sub FolderNames()
'Update 20141027
Application.ScreenUpdating = False
Dim xPath As String
Dim xWs As Worksheet
Dim fso As Object, j As Long, folder1 As Object
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Choose the folder"
    .Show
End With
On Error Resume Next
xPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"
Application.Workbooks.Add
Set xWs = Application.ActiveSheet
xWs.Cells(1, 1).Value = xPath
xWs.Cells(2, 1).Resize(1, 5).Value = Array("Path", "Dir", "Name", "Date Created", "Date Last Modified")
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder1 = fso.getFolder(xPath)
getSubFolder folder1
xWs.Cells(2, 1).Resize(1, 5).Interior.Color = 65535
xWs.Cells(2, 1).Resize(1, 5).EntireColumn.AutoFit
Application.ScreenUpdating = True
End Sub
Sub getSubFolder(ByRef prntfld As Object)
Dim SubFolder As Object
Dim subfld As Object
Dim xRow As Long
For Each SubFolder In prntfld.SubFolders
    xRow = Range("A1").End(xlDown).Row + 1
    Cells(xRow, 1).Resize(1, 5).Value = Array(SubFolder.Path, Left(SubFolder.Path, InStrRev(SubFolder.Path, "\")), SubFolder.Name, SubFolder.DateCreated, SubFolder.DateLastModified)
Next SubFolder
For Each subfld In prntfld.SubFolders
    getSubFolder subfld
Next subfld
End Sub

