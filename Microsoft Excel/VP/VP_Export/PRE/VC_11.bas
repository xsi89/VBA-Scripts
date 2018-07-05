Attribute VB_Name = "Volvo_Penta11"
Sub ColorRow_Blue_ONE()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Rows(1).Interior.Color = vbBlue
    Rows(1).Font.Color = vbRed


Next ws
 


End Sub


Sub Nonecolorcells_yel_run_Two()

    For Each sht In ActiveWorkbook.Worksheets
    Set Rng = sht.UsedRange
    Set MyRange = Rng
    
        For Each MyCol In MyRange.Columns
            For Each mycell In MyCol.Cells
                If mycell.Interior.ColorIndex = xlNone Then
                mycell.Interior.ColorIndex = 9
               
               End If
               
            Next
        Next
    Next

End Sub



Sub Splitbook_THREE()

Dim xPath As String
xPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each xWs In ThisWorkbook.Sheets
xWs.Copy
Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & xWs.Name & ".xls"
Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub






Sub LoopAllExcelFilesInFolder_FOUR()

'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myPath As String
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
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'gör något
Call Findcolumn_one
Nonecolorcells_yel_run_Two
Call clearcolor
'Call version_1
    
     
    ' MsgBox ActiveWorkbook.Name
     
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub



Sub Findcolumn_one()

    For Each sht In ActiveWorkbook.Worksheets
    Set Rng = sht.UsedRange
    Set MyRange = Rng

        For Each MyCol In MyRange.Columns
            For Each mycell In MyCol.Cells
                If mycell.Interior.ColorIndex = 23 Then
                    mycell.Font.ColorIndex = 1
                    Cells(mycell.Row, 2).Copy
                    mycell.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                End If
            
            Next
        Next
    Next
   
End Sub

Sub version_1()



myPath = "C:\"
MFileN = ActiveWorkbook.Name
myfilename = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)



     For Each sht In ActiveWorkbook.Worksheets
    Set Rng = sht.UsedRange
    Set MyRange = Rng
    
        For Each MyCol In MyRange.Columns
            For Each mycell In MyCol.Cells
                If mycell.Font.Color = vbRed Then
                  
                  'ActiveCell.EntireColumn.Copy
                 
                 If IsNumeric(mycell.Value) = False And _
                    IsError(mycell.Value) = False Then
                     
                     
                Cells(1, mycell.Column).Select
                ActiveCell.EntireColumn.Select
                ActiveCell.EntireColumn.Copy



  myName = MyCol.Cells(1, 1).Text



       NewWBName = ActiveWorkbook.Name

    Dim wbNew  As Workbook
    Dim wSheet As Worksheet
    Dim iSheet As Integer

   

    Set wbNew = Workbooks.Add
    iSheet = wbNew.Sheets.Count
    With wbNew
        For Each wSheet In ThisWorkbook.Sheets
            wSheet.Copy After:=.Sheets(.Sheets.Count)
        Next wSheet
    End With


      ActiveWorkbook.SaveAs Filename:=(myPath) & (myfilename) & "_" & (myName) & ".xls"


                
                
                 End If

                    End If
             
            Next
        Next
        Next

End Sub



Sub clearcolor()


    For Each sht In ActiveWorkbook.Worksheets
    Set Rng = sht.UsedRange
    Set MyRange = Rng
    
        For Each MyCol In MyRange.Columns
            For Each mycell In MyCol.Cells
                If mycell.Interior.ColorIndex = 9 Then mycell.ClearContents
                mycell.Interior.ColorIndex = xlNone
                
                
            Next
        Next
    Next


End Sub



Sub ColoRest()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Rows(1).Interior.Color = xlNone
    Rows(1).Font.ColorIndex = 1


Next ws
 


End Sub


'
'Sub Volvo_Penta()
'
'' Följande skript körs
'    Call Findcolumn_one ' detta letar efter alla blåa celler och gör texten röda sedan klistrar in dom i rätt kolumn resp (språk)
'    Call clearContent_two ' Detta skript går igenom alla öppna sheets, sedan så gör första radens bg = blå och även texten = röd
'    Call RemovecoloredRows_three ' detta skript letar efter grön bg och även texten och tar bort
'    Call RemovecoloredRows_four  ' detta skript letar efter grön bg och även texten och tar bort
'End Sub
'
'
'
''Sub clearContent_three()
'
'    For Each sht In ActiveWorkbook.Worksheets
'    Set Rng = sht.UsedRange
'    Set MyRange = Rng
'
'        For Each MyCol In MyRange.Columns
'            For Each mycell In MyCol.Cells
'                If mycell.Interior.ColorIndex = xlNone Then mycell.ClearContents
'                mycell.Font.ColorIndex = 1
'
'            Next
'        Next
'    Next
'
'End Sub
'
''Sub RemovecoloredRows_four()
'
'    For Each sht In ActiveWorkbook.Worksheets
'    Set Rng = sht.UsedRange
'    Set MyRange = Rng
'
'        For Each MyCol In MyRange.Columns
'            For Each mycell In MyCol.Cells
'                If mycell.Interior.ColorIndex = 3 Then mycell.ClearContents
'                    If mycell.Value = 0 Then
'                    mycell.Interior.ColorIndex = xlNone
'                    End If
'            Next
'        Next
'    Next
'
'End Sub
'
''Sub RemovecoloredRows_five()
'    For Each sht In ActiveWorkbook.Worksheets
'    Set Rng = sht.UsedRange
'    Set MyRange = Rng
'
'        For Each MyCol In MyRange.Columns
'            For Each mycell In MyCol.Cells
'                If mycell.Interior.ColorIndex = 4 Then mycell.ClearContents
'                    If mycell.Value = 0 Then
'                    mycell.Interior.ColorIndex = xlNone
'                    End If
'            Next
'        Next
'    Next
''End Sub
'
