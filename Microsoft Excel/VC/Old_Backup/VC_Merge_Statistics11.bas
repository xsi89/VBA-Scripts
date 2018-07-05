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

lastrow2 = Range("B" & Rows.Count).End(xlUp).row

With Worksheets("Volvo_Statistik")
num_rows = .Range("A1").End(xlDown).row

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
For I = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    If Cells(I, "J") = "IND" And Cells(I, "K") = "IND" And Len(Cells(I, "B")) <> 0 And Cells(I, "AB") <> 0 Then
    Cells(I, 1).EntireRow.Interior.ColorIndex = 6
    mycell = Cells(I, "B").Text
    mycellRes = mycell - 1
    DivAB = Cells(I, "AB").Text
    myTot = Round(DivAB / mycellRes)
    ValueSought = Cells(I, "A").Value
    Set FirstInstance = Columns("A").Find(what:=ValueSought, LookIn:=xlFormulas, LookAt:=xlWhole, SearchFormat:=False)
        If Not FirstInstance Is Nothing Then Set RangeToSearch = Range(FirstInstance, Cells.SpecialCells(xlCellTypeLastCell)).Columns(1)
        RangeToSearchFirstRow = FirstInstance.row
        RangeToSearchValues = RangeToSearch.Value
        Set RangeToUpdate = FirstInstance
        For j = 1 To UBound(RangeToSearchValues)
        If RangeToSearchValues(j, 1) = ValueSought Then
        Set RangeToUpdate = Union(RangeToUpdate, Cells(j - 1 + RangeToSearchFirstRow, "A"))
        End If
    Next j
    RangeToUpdate.Offset(, 19).Value = myTot
    End If
Next I
End Sub


Sub CheckMLY()

'MLY=MLY CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For I = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    If Cells(I, "J") = "MLY" And Cells(I, "K") = "MLY" And Len(Cells(I, "B")) <> 0 And Cells(I, "AB") <> 0 Then
    Cells(I, 1).EntireRow.Interior.ColorIndex = 4
    mycell = Cells(I, "B").Text
    mycellRes = mycell - 1
    DivAB = Cells(I, "AB").Text
    myTot = Round(DivAB / mycellRes)
    ValueSought = Cells(I, "A").Value
    Set FirstInstance = Columns("A").Find(what:=ValueSought, LookIn:=xlFormulas, LookAt:=xlWhole, SearchFormat:=False)
        If Not FirstInstance Is Nothing Then Set RangeToSearch = Range(FirstInstance, Cells.SpecialCells(xlCellTypeLastCell)).Columns(1)
        RangeToSearchFirstRow = FirstInstance.row
        RangeToSearchValues = RangeToSearch.Value
        Set RangeToUpdate = FirstInstance
        For j = 1 To UBound(RangeToSearchValues)
        If RangeToSearchValues(j, 1) = ValueSought Then
        Set RangeToUpdate = Union(RangeToUpdate, Cells(j - 1 + RangeToSearchFirstRow, "A"))
        End If
    Next j
    RangeToUpdate.Offset(, 19).Value = myTot
    End If
Next I
End Sub


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
            
            
                For I = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
                    mymonth = Val(Right(myyearandMonth, 2)) ' Remove the leading zero
                    thisMonth = mymonth
                    myCorrMonth = MonthName(thisMonth, True)
                    Cells(I, "G") = myyear
                    Cells(I, "H") = myCorrMonth
                    Cells(I, "D") = myCompany
                Next I
            
            End If
    Next sht
    
    
Application.DisplayAlerts = False
Rows("1:1").Select
Selection.EntireRow.Delete
Application.DisplayAlerts = True



Columns("B:B").Select
Selection.ClearContents
    
End Sub




Sub CountDiff()
ActiveWS = ActiveSheet.UsedRange.Rows.Count
Dim cell As Range

'No Match: 1,35
'Fuzzy: 0,67
'100/Rep.: 0,34
'DTP: 0,30/ord
'Internal h: 400,00
'Mini: 400,00

' Clear Column B


'----------------------

    For Each cell In Range("N1:N" & ActiveWS)
        If InStr(1, cell.Value, "", vbTextCompare) > 0 Then
        
            'No Match: 1,35
            IDNoMatch = Range("L" & cell.row).Value
            IDNoMatchRes = (IDNoMatch * 1.35)
            '--------------------------------------
            
            'Fuzzy: 0,67
            IDFuzzy = Range("M" & cell.row).Value
            IDFuzzyRes = (IDFuzzy * 0.67)
            '--------------------------------------
            
            '100/Rep.: 0,34
            IDRep = Range("N" & cell.row).Value
            IDRepRes = (IDRep * 0.34)
            '--------------------------------------
            
            ' Result of NoMatch + Fuzzy + 100 & Reps
            IDresultNFR = IDNoMatchRes + IDLFuzzyRes + IDRepRes
            '  MsgBox "NoMatch Words: " & IDNoMatchRes & vbNewLine & "Fuzzy Words: " & IDFuzzyRes & vbNewLine & "100 & Reps Words: " & IDRepRes
            
            IDP = Range("P" & cell.row).Value
            IDQ = Range("Q" & cell.row).Value
            IDR = Range("R" & cell.row).Value
            IDResultPQR = IDP + IDQ + IDR
            IDFinalResult = Round(IDresultNFR - IDResultPQR, 2)
            
            Range("C" & cell.row).Value = IDFinalResult 'result of math formula + cellvalue also do add in column P
            
            'MsgBox "No Match Words " & IDNoMatchRes & vbNewLine & "Fuzzy Matches: " & IDFuzzyRes & vbNewLine & "100% & Reps Words: " & IDRepRes & vbNewLine & "Total: " & IDresultNFR & vbNewLine & vbNewLine & "P: " & IDP & vbNewLine & "Q: " & IDQ & vbNewLine & "R: " & IDR & vbNewLine & "Q,P,R Total: " & IDResultPQR & vbNewLine & vbNewLine & "QPR -Matches :" & IDFinalResult
            
        End If
    Next

End Sub






Sub RemIND()

'RemoveINDD AND INDD;;;;;;;;;;;;;;;;;;;;;;
For I = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    If Cells(I, "J") = "IND" Then
    Cells(I, 1).EntireRow.Delete
    End If
Next I
End Sub


Sub RemMLY()

'Remove MLY ROW;;;;;;;;;;;;;;;;;;;;;;
For I = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    If Cells(I, "J") = "MLY" Then
    Cells(I, 1).EntireRow.Delete
    End If
Next I
End Sub



'RUN 7 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Sub Clearall()

'clear all color with this script

Cells.Interior.ColorIndex = xlNone


End Sub
'RUN 8 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



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




