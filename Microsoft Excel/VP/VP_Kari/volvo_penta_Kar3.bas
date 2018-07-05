Attribute VB_Name = "volvo_penta_Kari"


Sub Splitsheets()
'RUN 1
Dim XPath As String

myOrgName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))

XPath = Application.ActiveWorkbook.Path

MkDir XPath & "\FileSheets"
myLangPath = XPath & "\FileSheets"

Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each xWs In ThisWorkbook.Sheets
xWs.Copy
Application.ActiveWorkbook.SaveAs fileName:=myLangPath & "\" & myOrgName & "_" & xWs.Name & ".xls"
Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


 
 
Sub Batcher()
'RUN 2
Dim wb As Workbook
Dim myPath As String
Dim myfile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

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
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(fileName:=myPath & myfile)
    
    'gör något
Call CopyBtoC

    ' MsgBox ActiveWorkbook.Name
     
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      myfile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
 
 
 
 Sub CopyBtoC()

            
Rows("1:1").Select
Selection.EntireRow.Hidden = True
    For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
        Cells(i, 2).Copy
        Cells(i, 3).PasteSpecial xlPasteAll
        Columns("A:A").Select
        Selection.EntireColumn.Hidden = True
        Columns("B:B").Select
        Selection.EntireColumn.Hidden = True
    Next i
End Sub
            




