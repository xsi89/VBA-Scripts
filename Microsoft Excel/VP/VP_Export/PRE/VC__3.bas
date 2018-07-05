Attribute VB_Name = "Volvo_Penta_2"


Sub tesdewd()
Dim xPath As String
xPath = Application.ActiveWorkbook.Path
 
    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
                If myCell.Font.Color = vbRed Then
                  
                  'ActiveCell.EntireColumn.Copy
                 
                 If IsNumeric(myCell.Value) = False And _
                    IsError(myCell.Value) = False Then
                     
                     
                Cells(1, myCell.Column).Select
ActiveCell.EntireColumn.Select
ActiveCell.EntireColumn.Copy
                  
                  
                  
                  'MsgBox MyCell.Text
                  
                 End If
                    

                    End If
                    
                    
            Next
        Next
        Next

 
 
End Sub




Sub ColorRow1()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Rows(1).Interior.Color = vbBlue
    Rows(1).Font.Color = vbRed


Next ws
 


End Sub
   
    
  
       


Sub Splitbook()

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



Sub LoopAllExcelFilesInFolder()

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
    
    'g�r n�got
     
     MsgBox ActiveWorkbook.Name
     
    
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














































Sub Findcolumn_two()

    For Each sht In ActiveWorkbook.Worksheets
        Set rng = sht.UsedRange
       
    
    Next sht
    
    
End Sub






Sub Volvo_Penta()

' F�ljande skript k�rs
    Call Findcolumn_one ' detta letar efter alla bl�a celler och g�r texten r�da sedan klistrar in dom i r�tt kolumn resp (spr�k)
    Call clearContent_two ' Detta skript g�r igenom alla �ppna sheets, sedan s� g�r f�rsta radens bg = bl� och �ven texten = r�d
    Call RemovecoloredRows_three ' detta skript letar efter gr�n bg och �ven texten och tar bort
    Call RemovecoloredRows_four  ' detta skript letar efter gr�n bg och �ven texten och tar bort
End Sub

Sub Findcolumn_one()

    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng

        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
                If myCell.Interior.ColorIndex = 23 Then
                    myCell.Font.ColorIndex = 3
                    Cells(myCell.row, 2).Copy
                    myCell.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                End If
            
            Next
        Next
    Next
   
End Sub



'Sub clearContent_three()

    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
                If myCell.Interior.ColorIndex = xlNone Then myCell.ClearContents
                myCell.Font.ColorIndex = 1
               
            Next
        Next
    Next

End Sub

'Sub RemovecoloredRows_four()
    
    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
                If myCell.Interior.ColorIndex = 3 Then myCell.ClearContents
                    If myCell.Value = 0 Then
                    myCell.Interior.ColorIndex = xlNone
                    End If
            Next
        Next
    Next

End Sub

'Sub RemovecoloredRows_five()
    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
                If myCell.Interior.ColorIndex = 4 Then myCell.ClearContents
                    If myCell.Value = 0 Then
                    myCell.Interior.ColorIndex = xlNone
                    End If
            Next
        Next
    Next
'End Sub

