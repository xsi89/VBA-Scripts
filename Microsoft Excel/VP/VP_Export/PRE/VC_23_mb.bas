Attribute VB_Name = "Volvo_Penta20"
Sub ColorRow_Blue_ONE()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Rows(1).Interior.Color = vbBlue
    Rows(1).Font.Color = vbRed

Next ws
End Sub

Sub Splitbook_FOUR()

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


Sub LoopAllExcelFilesInFolder_FIVE()

Dim wb As Workbook
Dim myPath As String
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
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls"

'Target Path with Ending Extention
  MyFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While MyFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & MyFile)
    
    'gör något
Call ColorRow_Blue_ONE
Call Splitbook_FOUR
Call Findcolumn_one
Call Nonecolorcells_yel_run_Two
Call Redcolorcells_yel_run_THREE
Call clearcolor
Call clearcolumn_A_B
Call SaveColumns
Call ColoRest
    
     
    ' MsgBox ActiveWorkbook.Name
     
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      MyFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Sub Nonecolorcells_yel_run_Two()

    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each MyCell In MyCol.Cells
                If MyCell.Interior.ColorIndex = xlNone Then
                MyCell.Interior.ColorIndex = 9
               End If
               
            Next
        Next
    Next

End Sub

Sub Yellcolorcells_yel_run_Two()
'hitta gul
    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each MyCell In MyCol.Cells
                If MyCell.Interior.ColorIndex = 6 Then
                MyCell.Interior.ColorIndex = 9
               End If
               
            Next
        Next
    Next

End Sub

Sub Redcolorcells_yel_run_THREE()
'hitta röd
    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
        For Each MyCol In MyRange.Columns
            For Each MyCell In MyCol.Cells
                If MyCell.Interior.ColorIndex = 3 Then
                MyCell.Interior.ColorIndex = 9
                End If
            Next
        Next
    Next

End Sub

Sub Findcolumn_one()

    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng

        For Each MyCol In MyRange.Columns
            For Each MyCell In MyCol.Cells
                If MyCell.Interior.ColorIndex = 23 Then
                    MyCell.Font.ColorIndex = 1
                    Cells(MyCell.Row, 2).Copy
                    MyCell.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                End If
            
            Next
        Next
    Next
   
End Sub

Sub clearcolor()

    For Each sht In ActiveWorkbook.Worksheets
        Set rng = sht.UsedRange
        Set MyRange = rng
        
            For Each MyCol In MyRange.Columns
                For Each MyCell In MyCol.Cells
                If MyCell.Interior.ColorIndex = 9 Then MyCell.ClearContents
                MyCell.Interior.ColorIndex = xlNone
            
            Next
        Next
    Next
End Sub

Sub clearcolumn_A_B()
Columns("A:A").Select
Selection.ClearContents
Columns("B:B").Select
Selection.ClearContents

End Sub



Sub SaveColumns()
' sparar ner alla kolumner som seperat filer
    Dim wbNew As Workbook
    Dim wsSrc As Worksheet
    Dim cl As Range
    Dim rng As Range
    Mfilen = ActiveWorkbook.Name
    Set wsSrc = ActiveSheet    ' change as needed
    Set rng = wsSrc.UsedRange
        For Each cl In rng.Columns
            If cl.Cells(1, 1).Value <> "" Then
                Set wbNew = Workbooks.Add(xlWBATWorksheet)
                cl.Copy wbNew.Sheets(1).Range("A1")
                wbNew.SaveAs ThisWorkbook.Path & "\" & Mfilen & "_" & cl.Cells(1, 1).Value & ".xls", xlExcel8
                wbNew.Close
            End If
        Next cl
End Sub

Sub ColoRest()
'återställer färgen på row 1
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Rows(1).Interior.Color = xlNone
        Rows(1).Font.ColorIndex = 1
    Next ws
End Sub
