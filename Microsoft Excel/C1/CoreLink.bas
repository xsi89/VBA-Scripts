Attribute VB_Name = "CoreLink"
Sub ChangeFontColorBatch()
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
  myExtension = "*.xlsx"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myfile)

' Program börjar
            For Each Sht In ActiveWorkbook.Worksheets
            Set Rng = Sht.UsedRange
            
            Set myRange = Rng
            For Each mycol In myRange.Columns
            For Each mycell In mycol.Cells
            
            If mycell.Interior.ColorIndex = xlNone Then
            
                If mycell.Value <> "" Then
                mycell.Font.Color = RGB(255, 0, 0)
                'MsgBox mycell.Value
                ' Create a sheet with a name for each
                End If
            End If
            
            
            If mycell.Interior.ColorIndex = 2 Then
            If mycell.Value <> "" Then
            
            mycell.Font.Color = RGB(255, 0, 0)
            
            End If
            End If
            
            
            
            Next
            Next
            Next
            
            
'SLUT av programsats

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







Sub ChangeFontColor()


For Each Sht In ActiveWorkbook.Worksheets
Set Rng = Sht.UsedRange

Set myRange = Rng
For Each mycol In myRange.Columns
For Each mycell In mycol.Cells

If mycell.Interior.ColorIndex = xlNone Then

    If mycell.Value <> "" Then
    mycell.Font.Color = RGB(255, 0, 0)
    'MsgBox mycell.Value
    ' Create a sheet with a name for each
    End If
End If


If mycell.Interior.ColorIndex = 2 Then
If mycell.Value <> "" Then

mycell.Font.Color = RGB(255, 0, 0)

End If
End If



Next
Next
Next

End Sub






Sub ChangeFontBackBatch()
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
  myExtension = "*.xlsx"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myfile)


For Each Sht In ActiveWorkbook.Worksheets
Set Rng = Sht.UsedRange

Set myRange = Rng
For Each mycol In myRange.Columns
For Each mycell In mycol.Cells

If mycell.Interior.ColorIndex = xlNone Then


If mycell.Font.Color = RGB(255, 0, 0) Then


mycell.Font.Color = RGB(0, 0, 0)


End If

End If
Next
Next
Next

';;;;;;;;;;;;;;WRITE YOUR CODE HERE

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


Sub ChangeBack()


For Each Sht In ActiveWorkbook.Worksheets
Set Rng = Sht.UsedRange

Set myRange = Rng
For Each mycol In myRange.Columns
For Each mycell In mycol.Cells

If mycell.Interior.ColorIndex = xlNone Then


If mycell.Font.Color = RGB(255, 0, 0) Then


mycell.Font.Color = RGB(0, 0, 0)


End If

End If
Next
Next
Next

End Sub





