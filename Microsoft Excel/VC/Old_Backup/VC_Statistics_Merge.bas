Attribute VB_Name = "Volvo_Statistics_Merge"
':::::::::::::::::VOLVO STATISTICS MERGE:::::::::::::::::'






Sub Volvo_()

Dim StrCurrentfile As String
Dim intResult As Integer

Application.DisplayAlerts = True


intResult = Application.FileDialog(msoFileDialogFolderPicker).Show

If intResult = 0 Then

    MsgBox "User pressed cancel macro will stop!"

Exit Sub

Else

strDocPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"

End If

StrCurrentfile = Dir(strDocPath & "*NoTrans.xls")
Do While StrCurrentfile <> ""








Application.DisplayAlerts = False
StrCurrentfile.CheckCompatibility = False


StrCurrentfile = Dir


Loop
End Sub



