Attribute VB_Name = "Volvo_Penta_expTurk0311"
Sub Run_One()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Rows(1).Interior.Color = vbBlue
    Rows(1).Font.Color = vbRed

Next ws
End Sub

Sub Run_Two()
'Split book
Dim xPath As String

myOrgName = ThisWorkbook.Name

xPath = Application.ActiveWorkbook.Path

MkDir xPath & "\FileSheets"
MkDir xPath & "\LangCombNoT"
myLangPath = xPath & "\FileSheets"


 myFilenoEx = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))



Application.ScreenUpdating = False
Application.DisplayAlerts = False
    For Each xWs In ThisWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs Filename:=myLangPath & "\" & myFilenoEx & "_" & xWs.Name & ".xls"
    Application.ActiveWorkbook.Close False
    Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


Sub Run_Three()

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
Call alignment
Call NoneBG_Cells
'Call clearcolor
'Call clearContent
'Call clearcolumn_A_B
Call SaveColumns
'Call ColoRest
    
     
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


Sub Run_Four()
Application.DisplayAlerts = False

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
'  myExtension = "*_Tur.xls"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'g�r n�got

    Rows("1:1").Select
    Selection.EntireRow.Hidden = True
     
    ' MsgBox ActiveWorkbook.Name
     
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Nu �r alla celler p� Rad 1 Dolda!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Sub alignment()
' alignment2 Macro
    Range("A1:L817").Select
    Range("B20").Activate
    ActiveWindow.ScrollColumn = 2
    Columns("B:AA").Select
    Range("B20").Activate
    ActiveWindow.SmallScroll ToRight:=28
    Columns("B:BD").Select
    Range("B20").Activate
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlLTR
        .MergeCells = False
    End With
End Sub


Sub NoneBG_Cells()

    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
    Columns("A:A").Select
Selection.ClearContents
Columns("B:B").Select
Selection.ClearContents
                      
        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
            

            
                If myCell.Interior.ColorIndex = xlNone Then
               myCell.Interior.ColorIndex = 3
                
               End If
               
               If myCell.Value = "" Then
   myCell.Interior.ColorIndex = 8
         
            
            End If
            
            If myCell.Interior.Color = 3 Then
            myCell.Interior.ColorIndex = xlNone
                End If
               
            Next
        Next
    Next

End Sub





Sub SaveColumns()
Application.DisplayAlerts = False




    Dim wbNew As Workbook
    Dim wsSrc As Worksheet
    Dim cl As Range
    Dim rng As Range
    MFileN = ActiveWorkbook.Name
    
    
   NewName = Left(MFileN, (InStrRev(MFileN, ".", -1, vbTextCompare) - 1))
  
    Set wsSrc = ActiveSheet    ' change as needed
    Set rng = wsSrc.UsedRange
        For Each cl In rng.Columns
            If cl.Cells(1, 1).Value <> "" Then
                Set wbNew = Workbooks.Add(xlWBATWorksheet)
                cl.Copy wbNew.Sheets(1).Range("A1")
                Sheets(1).Name = "WordNotTrans"
                wbNew.CheckCompatibility = False
               'wbNew.SaveAs ThisWorkbook.Path & "\" & NewName & "_" & cl.Cells(1, 1)
                wbNew.SaveAs ThisWorkbook.Path & "\" & NewName & "_" & cl.Cells(1, 1) & "_NoTrans" & ".xls", xlExcel8
                wbNew.Close
            End If
        Next cl
End Sub

Sub ColoRest()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Rows(1).Interior.Color = xlNone
        Rows(1).Font.ColorIndex = 1
    Next ws
End Sub











































