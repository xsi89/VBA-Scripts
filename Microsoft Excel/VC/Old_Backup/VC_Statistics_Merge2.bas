Attribute VB_Name = "Volvo_Merge_Statistics"
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

StrCurrentfile = Dir(strDocPath & "*.xls")
Do While StrCurrentfile <> ""







Application.DisplayAlerts = False
myLangFileN.Close SaveChanges:=False
noLangFilen.CheckCompatibility = False
noLangFilen.Close SaveChanges:=True

StrCurrentfile = Dir


Loop


End Sub


Sub mypathen()


Dim folderPath As String
folderPath = Application.ActiveWorkbook.Path


End Sub

Sub Merge2MultiSheets()
Dim wbDst As Workbook
Dim wbSrc As Workbook
Dim wsSrc As Worksheet
Dim MyPath As String
Dim strFilename As String
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    MyPath = "C:\MyPath" ' change to suit
    Set wbDst = Workbooks.Add(xlWBATWorksheet)
    strFilename = Dir(MyPath & "\*.xls", vbNormal)
    
    If Len(strFilename) = 0 Then Exit Sub
    
    Do Until strFilename = ""
        
            Set wbSrc = Workbooks.Open(Filename:=MyPath & "\" & strFilename)
            
            Set wsSrc = wbSrc.Worksheets(1)
            
            wsSrc.Copy After:=wbDst.Worksheets(wbDst.Worksheets.Count)
            
            wbSrc.Close False
        
        strFilename = Dir()
        
    Loop
    wbDst.Worksheets(1).Delete
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub
