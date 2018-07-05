Attribute VB_Name = "Volvo_Penta_import"
Sub Run_Four()
Application.DisplayAlerts = False

Dim wb As Workbook
Dim MyPath As String
Dim myFile As String
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
  myFile = Dir(MyPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=MyPath & myFile)
    
    'gör något







    
    'Save and Close Workbook
      wb.Close savechanges:=True

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Nu är dom mergade"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub



'The following code will combine all data into one excel workbook.
Sub CombineFiles_Step1()
'Declare Variables
Dim WorkbookDestination As Workbook
Dim WorkbookSource As Workbook
Dim WorksheetSource As Worksheet
Dim FolderLocation As String
Dim strFilename As String
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'This line will need to be modified depending on location of source folder
    FolderLocation = "C:\Users\daniel.elmnas.TT\Desktop\ko\FIle"
    
    'Set the current directory to the the folder path.
    ChDrive FolderLocation
    ChDir FolderLocation
    
    'Dialog box to determine which files to use. Use ctrl+a to select all files in folder.
    SelectedFiles = Application.GetOpenFilename( _
        filefilter:="Excel Files (*.xls*), *.xls*", MultiSelect:=True)
    
    'Create a new workbook
    Set WorkbookDestination = Workbooks.Add(xlWBATWorksheet)
    strFilename = Dir(FolderLocation & "\*.xls", vbNormal)
    
    'Iterate for each file in folder
    If Len(strFilename) = 0 Then Exit Sub
    
    
    Do Until strFilename = ""
        
            Set WorkbookSource = Workbooks.Open(Filename:=FolderLocation & "\" & strFilename)
            Set WorksheetSource = WorkbookSource.Worksheets(1)
            WorksheetSource.Copy After:=WorkbookDestination.Worksheets(WorkbookDestination.Worksheets.Count)
            WorkbookSource.Close False
        strFilename = Dir()
        
    Loop
    WorkbookDestination.Worksheets(1).Delete
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
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







Sub App_FileSearch_Example()

Workbooks.Open Filename:="C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org - Copy\_*&*.xls"

End Sub




'The following code will combine all data into one excel workbook.
Sub Combinles_Step1()
'Declare Variables
Dim WorkbookDestination As Workbook
Dim WorkbookSource As Workbook
Dim WorksheetSource As Worksheet
Dim FolderLocation As String
Dim strFilename As String
    
    
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
                filefilter:="Excel Files (*.xls*), *.xls*", MultiSelect:=True)
            
            'Create a new workbook
            Set WorkbookDestination = Workbooks.Add(xlWBATWorksheet)
            strFilename = Dir(FolderLocation & "\*.xls", vbNormal)
            
            'Iterate for each file in folder
            If Len(strFilename) = 0 Then Exit Sub
            
            
            Do Until strFilename = ""
                
                    Set WorkbookSource = Workbooks.Open(Filename:=FolderLocation & "\" & strFilename)
                    Set WorksheetSource = WorkbookSource.Worksheets(1)
                    WorksheetSource.Copy After:=WorkbookDestination.Worksheets(WorkbookDestination.Worksheets.Count)
                    WorkbookSource.Close False
                strFilename = Dir()
                
            Loop
            WorkbookDestination.Worksheets(1).Delete
            
            Application.DisplayAlerts = True
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        End If
    End With
End Sub

