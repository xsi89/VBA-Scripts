Attribute VB_Name = "Volvo_Merge_Statistics"

Sub mypathen()


 MsgBox ThisWorkbook.Name & Chr(10) & ThisWorkbook.Path & Chr(10) & ActiveSheet.Name



End Sub

'(RUN 1:::::::::::::)
'Run this Code to Clean up the resource also orginal document from such sheetnames like:(Sheet1,sheet2 etc)

Sub DeleteSheets()

Dim wb As Workbook
Dim myPath As String
Dim myfile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim XPath As String
XPath = Application.ActiveWorkbook.Path


'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPat
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(FileName:=myPath & myfile)
    
    'gör något
Application.DisplayAlerts = False
    Dim sh As Worksheet
    For Each sh In Sheets
        If IsEmpty(sh.UsedRange) Then sh.Delete
    Next


On Error Resume Next
Sheets("Orders").Delete
On Error GoTo 0

     
    ' MsgBox ActiveWorkbook.Name
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      myfile = Dir
  Loop

'Message Box when tasks are completed
MsgBox "Nu är sheets bortagna"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

'(RUN 2:::::::::::::)
Sub CombineFiles()
     
    Dim Path As String
    Dim FileName As String
    Dim Wkb As Workbook
    Dim ws As Worksheet
    

    intResult = Application.FileDialog(msoFileDialogFolderPicker).Show

If intResult = 0 Then

    MsgBox "User pressed cancel macro will stop!"

Exit Sub

Else

strDocPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"

End If
     
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    FileName = Dir(strDocPath & "\*.xls", vbNormal)
    Do Until FileName = ""
        Set Wkb = Workbooks.Open(FileName:=strDocPath & "\" & FileName)
        For Each ws In Wkb.Worksheets
        
        
         Application.DisplayAlerts = False
 
        
        wbname = Replace(FileName, ".xls", "")

       'MsgBox WBname
            ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            
        Next ws
        ActiveSheet.Name = (wbname)
        Wkb.Close False
        FileName = Dir()
    Loop
    
    
    Worksheets("Volvo_Row_one").Move Before:=Worksheets(1)
    
    
    
     Dim sh As Worksheet
    For Each sh In Sheets
        If IsEmpty(sh.UsedRange) Then sh.Delete
    Next


    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

  Sub WorksheetLoop2()
    
 
For Each ws In Worksheets


   myOutput As String
      
      
     myOutput = Left(ws.Name, InStrRev(ws.Name, ".") - 1)
      
      MsgBox myOutput
      
      
       ActiveSheet.Name = myOutput
  
      
      
Next ws

  


  


  
   

         

      End Sub







Sub Volvo_()

Dim StrCurrentfile As String
Dim intResult As Integer

Dim wb As Workbook


Set wb = ActiveWorkbook
Dim OWb As Workbook

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




Set OWb = Workbooks.Open(strDocPath & StrCurrentfile)


OWb.Activate


Dim myname
'myname = Replace(ActiveWorkbook.Name, ".xls", "")
    ActiveSheet.Select
   ' ActiveSheet.Name = myname
    ActiveSheet.Name = "StatRowOne"
OWb.Sheets("StatRowOne").Copy
'ActiveSheet.Copy


wb.Activate
ActiveSheet.Paste


Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
'ws.Name = "Tempo"





'Application.DisplayAlerts = False
'myLangFileN.Close SaveChanges:=False
'noLangFilen.CheckCompatibility = False
'noLangFilen.Close SaveChanges:=True

StrCurrentfile = Dir


Loop


End Sub
