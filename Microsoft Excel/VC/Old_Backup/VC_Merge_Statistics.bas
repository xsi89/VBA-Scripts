Attribute VB_Name = "Volvo_Merge_Statistics"
'(RUN 1:::::::::::::)
'Run this Code to Clean up the resource also orginal document from such sheetnames like:(Sheet1,sheet2 etc)
Sub Deletesheets()

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
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings
  
  myExtension = "*.xls"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(fileName:=myPath & myfile)
    
    'gör något

 Application.DisplayAlerts = False
 
 'Sheets("Orders").Delete

    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Orders").Delete

    On Error GoTo 0

 
    Dim Sh As Worksheet
    For Each Sh In Sheets
        If IsEmpty(Sh.UsedRange) Then Sh.Delete
    Next

Application.DisplayAlerts = True
    ' MsgBox ActiveWorkbook.Name
     
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      myfile = Dir
  Loop

'Message Box when tasks are completed
'MsgBox "Nu är Kolumnfiler generade i mappen: " & XPath

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

'(RUN 2:::::::::::::)
Sub CombineFiles()
     
    Dim Path As String
    Dim fileName As String
    Dim Wkb As Workbook
    Dim ws As Worksheet
    

ActiveSheet.Name = "Volvo_Statistik"
    intResult = Application.FileDialog(msoFileDialogFolderPicker).Show

If intResult = 0 Then

    MsgBox "User pressed cancel macro will stop!"

Exit Sub

Else

strDocPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"

End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    fileName = Dir(strDocPath & "\*.xls", vbNormal)
    Do Until fileName = ""
        Set Wkb = Workbooks.Open(fileName:=strDocPath & "\" & fileName)
        For Each ws In Wkb.Worksheets
         Application.DisplayAlerts = False
         wbname = Replace(fileName, ".xls", "")

       'MsgBox WBname
            ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            
        Next ws
        ActiveSheet.Name = (wbname)
        Wkb.Close False
        fileName = Dir()
    Loop
    
    
    Worksheets("Volvo_Statistik").Move Before:=Worksheets(1)

    
    
'     Dim sh As Worksheet
'    For Each sh In Sheets
'        If IsEmpty(sh.UsedRange) Then sh.Delete
'    Next


    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

'RUN 4;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Sub GetData()
    For Each sht In ThisWorkbook.Worksheets
    mysheetname = sht.Name
        If sht.Name Like "Volvo Penta*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.Paste
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.Paste
        Sheets(mysheetname).Select
        Columns("D:E").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        Columns("J:J").Select
        ActiveSheet.Paste
        Sheets(mysheetname).Select
         Columns("G:I").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.Paste
        Sheets(mysheetname).Select
         Columns("L:L").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("W:W").Select
         ActiveSheet.Paste
        Sheets(mysheetname).Select
        Columns("N:N").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.Paste
        End If
    Next sht
End Sub

'RUN 5;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
'IND CHECK FIL;;;;;;;;;;;;;;;;;;;;;;


Sub CountIDer()
Dim num_rows As Long

lastrow2 = Range("B" & Rows.Count).End(xlUp).Row

With Worksheets("Volvo_Statistik")
num_rows = .Range("A1").End(xlDown).Row

.Range("B2").Formula = "=IF(A2<>A3,COUNTIF(A$2:A2,A2),"""")"
.Range("B2:B" & num_rows).FillDown

 'I autofil all cells down"
'Range("B1").Select
'ActiveCell.FormulaR1C1 = "=COUNTIF(C[-1],RC[-1])"
'Selection.AutoFill Destination:=Range("B1:B" & lastrow2), Type:=xlFillDefault

End With
End Sub

Sub CheckIND()

'IND=IND CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "J") = "IND" And Cells(i, "K") = "IND" And Len(Cells(i, "B")) <> 0 And Cells(i, "AB") <> 0 Then
    Cells(i, 1).EntireRow.Interior.ColorIndex = 6
    mycell = Cells(i, "B").Text
    mycellRes = mycell - 1
    DivAB = Cells(i, "AB").Text
    myTot = Round(DivAB / mycellRes)
    ValueSought = Cells(i, "A").Value
    Set FirstInstance = Columns("A").Find(what:=ValueSought, LookIn:=xlFormulas, LookAt:=xlWhole, SearchFormat:=False)
        If Not FirstInstance Is Nothing Then Set RangeToSearch = Range(FirstInstance, Cells.SpecialCells(xlCellTypeLastCell)).Columns(1)
        RangeToSearchFirstRow = FirstInstance.Row
        RangeToSearchValues = RangeToSearch.Value
        Set RangeToUpdate = FirstInstance
        For j = 1 To UBound(RangeToSearchValues)
        If RangeToSearchValues(j, 1) = ValueSought Then
        Set RangeToUpdate = Union(RangeToUpdate, Cells(j - 1 + RangeToSearchFirstRow, "A"))
        End If
    Next j
    RangeToUpdate.Offset(, 19).Value = myTot
    End If
Next i
End Sub




Sub CheckMLY()

'MLY=MLY CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "J") = "MLY" And Cells(i, "K") = "MLY" And Len(Cells(i, "B")) <> 0 And Cells(i, "AB") <> 0 Then
    Cells(i, 1).EntireRow.Interior.ColorIndex = 4
    mycell = Cells(i, "B").Text
    mycellRes = mycell - 1
    DivAB = Cells(i, "AB").Text
    myTot = Round(DivAB / mycellRes)
    ValueSought = Cells(i, "A").Value
    Set FirstInstance = Columns("A").Find(what:=ValueSought, LookIn:=xlFormulas, LookAt:=xlWhole, SearchFormat:=False)
        If Not FirstInstance Is Nothing Then Set RangeToSearch = Range(FirstInstance, Cells.SpecialCells(xlCellTypeLastCell)).Columns(1)
        RangeToSearchFirstRow = FirstInstance.Row
        RangeToSearchValues = RangeToSearch.Value
        Set RangeToUpdate = FirstInstance
        For j = 1 To UBound(RangeToSearchValues)
        If RangeToSearchValues(j, 1) = ValueSought Then
        Set RangeToUpdate = Union(RangeToUpdate, Cells(j - 1 + RangeToSearchFirstRow, "A"))
        End If
    Next j
    RangeToUpdate.Offset(, 19).Value = myTot
    End If
Next i
End Sub






Sub RemIND()

'RemoveINDD AND INDD;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "J") = "IND" Then
    Cells(i, 1).EntireRow.Delete
    End If
Next i
End Sub


Sub RemMLY()

'Remove MLY ROW;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "J") = "MLY" Then
    Cells(i, 1).EntireRow.Delete
    End If
Next i
End Sub


'RUN 7 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Sub Clearall()

'clear all color with this script

Cells.Interior.ColorIndex = xlNone
Columns("B:B").Select
Selection.ClearContents

End Sub
'RUN 8 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Sub Companyname()

Dim mysheetname As String
Dim Myarr() As String
Dim myCompany As String
Dim myyearandMonth As String
Dim myyear As String
Dim mymonth As String
Dim myvar As String
Dim thisMonth As Integer
Dim myCorrMonth As String

    For Each sht In ThisWorkbook.Worksheets
        mysheetname = sht.Name
            If sht.Name Like "Volvo Penta*" Then
            
            Worksheets("Volvo_Statistik").Activate
            Myarr = Split(mysheetname, "_")
            myCompany = Myarr(0)
            myyearandMonth = Myarr(1)
            myyear = Left(myyearandMonth, 4)
            mymonth = Right(myyearandMonth, 2)
            
            
                For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
                    mymonth = Val(Right(myyearandMonth, 2)) ' Remove the leading zero
                    thisMonth = mymonth
                    myCorrMonth = MonthName(thisMonth, True)
                    Cells(i, "G") = myyear
                    Cells(i, "H") = myCorrMonth
                    Cells(i, "D") = myCompany
                Next i
            
            End If
    Next sht
    
End Sub


Sub SetFormsValCorr()

'RUN 9 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Call NoMatch
Call FuzzyMatch
Call HundredandREP

End Sub



Sub NoMatch()

ActiveWS = ActiveSheet.UsedRange.Rows.Count 'count the number of rows used'
Dim cell As Range

For Each cell In Range("L1:L" & ActiveWS) 'define range of cells
    If InStr(1, cell.Value, "", vbTextCompare) > 0 Then ' if zero
    myresult = Val(cell.Value * 1.35) ' math formula
    Range("P" & cell.Row).Value = myresult 'result of math formula + cellvalue also do add in column P
    End If
Next

End Sub


Sub FuzzyMatch()

ActiveWS = ActiveSheet.UsedRange.Rows.Count 'count the number of rows used'
Dim cell As Range

For Each cell In Range("M1:M" & ActiveWS) 'define range of cells
    If InStr(1, cell.Value, "", vbTextCompare) > 0 Then ' if zero
     myresult = Val(cell.Value * 0.67) ' math formula
    Range("Q" & cell.Row).Value = myresult 'result of math formula + cellvalue also do add in column P
    End If
Next

End Sub

Sub HundredandREP()

ActiveWS = ActiveSheet.UsedRange.Rows.Count 'count the number of rows used'
Dim cell As Range

For Each cell In Range("N1:N" & ActiveWS) 'define range of cells
    If InStr(1, cell.Value, "", vbTextCompare) > 0 Then ' if zero
    myresult = Val(cell.Value * 0.34) ' math formula
    Range("R" & cell.Row).Value = myresult 'result of math formula + cellvalue also do add in column P
    End If
Next

End Sub













Sub ListFilesInFolder()
     

'Dim sPath As String
sPath = "C:\test\"
'Shell "C:\WINDOWS\explorer.exe """ & sPath & "", vbNormalFocus


 Shell ("explorer.exe ""search-ms://query=34234""")


'Shell("c:\Windows\explorer.exe
'C:\Users\<username>\Searches\<searchname>.search-ms

End Sub



Sub myqeqwe()

searchwin = "34228"
   pathwin = "X:\1 Övriga kunder\Arkiv\Originalfiler"
Call Shell("explorer ""search-ms://query=" & searchwin & "&crumb=location:" & pathwin & """", vbNormalFocus)

   
   
End Sub




