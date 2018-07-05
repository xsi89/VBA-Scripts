Attribute VB_Name = "Volvo_Penta_import"
Sub Run_Four()
Application.DisplayAlerts = False

Dim wb As Workbook
Dim MyPath As String
Dim MyFile As String
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
        MyPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  MyPath = MyPath
  If MyPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls"

'Target Path with Ending Extention
  MyFile = Dir(MyPath & myExtension)

'Loop through each Excel file in folder
  Do While MyFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=MyPath & MyFile)
    
    'gör något







    
    'Save and Close Workbook
      wb.Close savechanges:=True

    'Get next file name
      MyFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Nu är dom mergade"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub




Sub Move_Colored_Cells()
Dim s As String
Dim Current As Worksheet
For Each Current In Worksheets
For Each c In ActiveSheet.UsedRange
If c.Interior.ColorIndex = 3 Then
s = c.Address
c.Copy Sheets("Translated").Range(s)
End If
Next
Next
End Sub






'The following code will combine all data into one excel workbook.
Sub FindFileName()
'Declare Variables
Dim WorkbookDestination As Workbook
Dim WorkbookSource As Workbook
Dim WorksheetSource As Worksheet
Dim FolderLocation As String
Dim strFileName As String
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .AllowMultiSelect = False
        .Title = "Select Source folder"
        If .Show = -1 Then
        
            Application.DisplayAlerts = False
            Application.EnableEvents = False
            Application.ScreenUpdating = False
        
            FolderLocation = .SelectedItems(1)
            

            
            'Dialog box to determine which files to use. Use ctrl+a to select all files in folder.
            SelectedFiles = Application.GetOpenFilename( _
                FileFilter:="Excel Files (*.xls*), *.xls*", MultiSelect:=True)
            
            'Create a new workbook
            Set WorkbookDestination = Workbooks.Add(xlWBATWorksheet)
            strFileName = Dir(FolderLocation & "\*.xls", vbNormal)
            
            'Iterate for each file in folder
            If Len(strFileName) = 0 Then Exit Sub
            
            
            Do Until strFileName = ""
                
                    Set WorkbookSource = Workbooks.Open(Filename:=FolderLocation & "\" & strFileName)
                    Set WorksheetSource = WorkbookSource.Worksheets(1)
                    WorksheetSource.Copy After:=WorkbookDestination.Worksheets(WorkbookDestination.Worksheets.Count)
                    WorkbookSource.Close False
                strFileName = Dir()
                
            Loop
            WorkbookDestination.Worksheets(1).Delete
            
            Application.DisplayAlerts = True
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        End If
    End With
End Sub























Sub files()
Application.DisplayAlerts = False

Dim wb As Workbook
Dim MyPath As String
Dim MyFile As String
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
        MyPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  MyPath = MyPath
  If MyPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*NoTrans.xls"

'Target Path with Ending Extention
  MyFile = Dir(MyPath & myExtension)

'Loop through each Excel file in folder
  Do While MyFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=MyPath & MyFile)
    
    'gör något
    
    
MyFileNoTransOpen = ActiveWorkbook.Name


myOpenNoTrans = Left(MyFileNoTransOpen, (InStrRev(MyFileNoTransOpen, ".", -1, vbTextCompare) - 9))



myLangFile = myOpenNoTrans & ".xls"


MsgBox myOpenNoTrans

MsgBox myLangFile








     
    ' MsgBox ActiveWorkbook.Name
     
    
    'Save and Close Workbook
      wb.Close savechanges:=True

    'Get next file name
      MyFile = Dir
  Loop

'Message Box when tasks are completed
 ' MsgBox "Nu är alla celler på Rad 1 Dolda!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
