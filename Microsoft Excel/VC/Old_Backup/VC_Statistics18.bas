Attribute VB_Name = "Volvo_Statistics"
'(RUN 1:::::::::::::)
'Run this Code to Clean up the resource also orginal document from such sheetnames like:(Sheet1,sheet2 etc)
Sub DelSheets()

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
  
  myExtension = "*.xlsx"

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
'This code Merging the Volvo resource files)
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

Sub DelUnSheets()
'That is fine if you get a error that says "intervall" something this is a  extra function to see if there are any empty unused sheets.
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Sheets("Blad1").Delete
ThisWorkbook.Sheets("Blad2").Delete
ThisWorkbook.Sheets("Blad3").Delete
ThisWorkbook.Sheets("Sheet1").Delete
ThisWorkbook.Sheets("Sheet2").Delete
ThisWorkbook.Sheets("Sheet3").Delete
On Error GoTo 0
Application.DisplayAlerts = True

End Sub

Sub GetColData()
'This function Get all the column data from the "Company" sheet.
' If name "xxx" then copy following sheets to current workbook.
For Each sht In ThisWorkbook.Worksheets
    mysheetname = sht.Name
    
        If sht.Name Like "Volvo_3P*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("D:D").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("D:D").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("H:H").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("G:G").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("J:J").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("H:H").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("I:I").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("J:J").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("M:M").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("E:E").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("M:M").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("U:U").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("O:O").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.paste
        End If
        
        If sht.Name Like "Volvo_Penta*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("D:D").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("D:D").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("H:H").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("G:G").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("J:J").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("H:H").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("I:I").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("J:J").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("M:M").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("E:E").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("M:M").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("U:U").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("O:O").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.paste
        End If
        
        If sht.Name Like "Volvo_Bus*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("D:D").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("D:D").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("H:H").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("G:G").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("J:J").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("H:H").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("I:I").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("J:J").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("M:M").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("E:E").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("M:M").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("U:U").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("O:O").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.paste
        End If
        
        If sht.Name Like "Volvo_Business_Service*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("D:D").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("D:D").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("H:H").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("G:G").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("J:J").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("H:H").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("I:I").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("J:J").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("M:M").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("E:E").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("M:M").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("U:U").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("O:O").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.paste
        End If
        
        If sht.Name Like "Volvo_Group_Sweden*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("D:D").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("D:D").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("H:H").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("G:G").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("J:J").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("H:H").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("I:I").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("J:J").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("M:M").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("E:E").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("M:M").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("U:U").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("O:O").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.paste
        End If
        
        If sht.Name Like "Volvo_IT*" Then
        Sheets(mysheetname).Select
        Columns("A:A").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("A:A").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("C:C").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("D:D").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("D:D").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("H:H").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("E:E").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("I:I").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("G:G").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("J:J").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("H:H").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("K:K").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("I:I").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("L:L").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("J:J").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("M:M").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("E:E").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("K:K").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("F:F").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("M:M").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("U:U").Select
        ActiveSheet.paste
        
        Sheets(mysheetname).Select
        Columns("O:O").Select
        Selection.Copy
        Sheets("Volvo_Statistik").Select
        Columns("AB:AB").Select
        ActiveSheet.paste
        End If
     
        
   
  
    Next sht
End Sub
Sub myyear()
'This function removes the 6 last characters from the word and then replace with the first 4 characters.Sheets("Volvo_Statistik").Activate
    For Each sht In ThisWorkbook.Worksheets
    mysheetname = sht.Name
    
        If sht.Name Like "Volvo_Statistik*" Then
        
        ActiveWS = ActiveSheet.UsedRange.Rows.Count
        Dim cell As Range
        Dim mystring As String
            For Each cell In Range("E1:E" & ActiveWS)
            myresult = Left(cell.Text, Len(cell.Text) - 6)
            cell.Value = myresult
            Next
            
        Else
        MsgBox "Sheetet Volvo_Statistik Does not exists"
        End If
    Exit For
    
    
    Next sht
    
End Sub



Sub MyMonthPone()
'This function removes

    For Each sht In ThisWorkbook.Worksheets
    mysheetname = sht.Name
    
        If sht.Name Like "Volvo_Statistik*" Then
        
        ActiveWS = ActiveSheet.UsedRange.Rows.Count
        Dim cell As Range
        Dim mystring As String
            
            For Each cell In Range("F1:F" & ActiveWS)
            myresult = Left(cell.Text, Len(cell.Text) - 3)
            
           '  MsgBox myresult
            cell.Value = myresult
            Next
        Else
        MsgBox "sheets:Volvo_Statistik doest not exists"
        End If
    Exit For
    Next sht

End Sub

Sub MyMonthPTwo()

    For Each sht In ThisWorkbook.Worksheets
    mysheetname = sht.Name
    
            If sht.Name Like "Volvo_Statistik*" Then
            
            ActiveWS = ActiveSheet.UsedRange.Rows.Count
            Dim cell As Range
            Dim mystring As String
            
                For Each cell In Range("F1:F" & ActiveWS)
                myresult = Right(cell.Text, Len(cell.Text) - 5)
                'MsgBox myresult
                cell.Value = "" & myresult
                Next
            Else
            MsgBox "sheets:Volvo_Statistik doest not exists"
            End If
        Exit For
    Next sht

End Sub

'Count how many repetitions there are of a ordernumber ;;;;;;;;;;;;;;;;;;;;;;
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
    
    'DELETE ROW ONE
    Rows(1).EntireRow.Delete
    
    
    End With
End Sub



Sub CheckIND()

'IND=IND CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "H") = "IND" And Cells(i, "I") = "IND" And Len(Cells(i, "B")) <> 0 And Cells(i, "AB") <> 0 Then
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
    RangeToUpdate.Offset(, 17).Value = myTot
    End If



Next i
End Sub



Sub CheckMLY()

'IND=IND CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "H") = "MLY" And Cells(i, "I") = "MLY" And Len(Cells(i, "B")) <> 0 And Cells(i, "AB") <> 0 Then
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
    RangeToUpdate.Offset(, 17).Value = myTot
    End If
Next i
End Sub



Sub RemIND()

'RemoveINDD AND INDD;;;;;;;;;;;;;;;;;;;;;;
For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "H") = "IND" Then
    Cells(i, 1).EntireRow.Delete
    End If
Next i
End Sub

Sub RemMLY()

'Remove MLY ROW;;;;;;;;;;;;;;;;;;;;;;
For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "H") = "MLY" Then
    Cells(i, 1).EntireRow.Delete
    End If
Next i
End Sub




Sub Company()

Dim wb As Workbook
Dim ws As Worksheet
Set wb = ActiveWorkbook

Columns(2).EntireColumn.ClearContents

        For Each ws In wb.Worksheets
            If ws.Name = "Volvo_3P" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo 3P"
            Next i
            End If
            
            If ws.Name = "Volvo_Penta" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo Penta"
            Next i
            End If
            
            If ws.Name = "Volvo_Bus" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo Bus"
            Next i
            End If
            
            If ws.Name = "Volvo_Business_Service" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo Business Service"
            Next i
            End If
            
            If ws.Name = "Volvo_Group_Sweden" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo Group Sweden"
            Next i
            End If
            
            If ws.Name = "Volvo_IT" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo IT"
            Next i
        End If
    Next
    
End Sub




Sub NewWordsC()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim n As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).Row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    n = ws2.Range("A" & ws2.Rows.Count).End(xlUp).Row
    
    For r = 1 To m
        For s = 1 To n
            If Trim(ws1.Range("H" & r) & ws1.Range("I" & r)) = Trim(ws2.Range("A" & s) & ws2.Range("B" & s)) Then
            
            
            
                 '   MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                    
                  
                                    
'                    Sheets("Volvo_Statistik").Activate
                   '  Range("J" & r).Select
                   ' MsgBox ws1.Range("J" & r).Value
'
'
'
                  '   Sheets("volvo_NewPrices").Activate
                   '  Range("G" & s).Select
                   '  MsgBox ws2.Range("G" & s).Value

                 ' ws1.Range("N" & r).Value = Val(ws1.Range("J" & r)) * Val(ws2.Range("G" & s))
                        
                 
                 
                   myStringRes = Val(ws1.Range("J" & r)) * Val(ws2.Range("G" & s))
                    myRes = Round(myStringRes, 2)
                    ws1.Range("N" & r).Value = "" & myRes
                 
                 
                 
            End If
        Next s
    Next r
End Sub



Sub FuzzyC()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim n As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).Row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    n = ws2.Range("A" & ws2.Rows.Count).End(xlUp).Row
    
    For r = 1 To m
        For s = 1 To n
            If Trim(ws1.Range("H" & r) & ws1.Range("I" & r)) = Trim(ws2.Range("A" & s) & ws2.Range("B" & s)) Then
            
            
            
                  '  MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                    
                  
                                    
''                    Sheets("Volvo_Statistik").Activate
'                     Range("J" & r).Select
'                     MsgBox ws1.Range("K" & r).Value
'
'                     Sheets("volvo_NewPrices").Activate
'                     Range("F" & s).Select
'                     MsgBox ws2.Range("F" & s).Value

                   myStringRes = Val(ws1.Range("K" & r)) * Val(ws2.Range("F" & s))
                    myRes = Round(myStringRes, 2)
                    ws1.Range("O" & r).Value = "" & myRes
                    
            End If
        Next s
    Next r
End Sub

Sub RepsandHund()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim n As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).Row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    n = ws2.Range("A" & ws2.Rows.Count).End(xlUp).Row
    
    For r = 1 To m
        For s = 1 To n
            If Trim(ws1.Range("H" & r) & ws1.Range("I" & r)) = Trim(ws2.Range("A" & s) & ws2.Range("B" & s)) Then
            
            
            
                  '  MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                    
                  
                                    
''                    Sheets("Volvo_Statistik").Activate
'                     Range("J" & r).Select
'                     MsgBox ws1.Range("K" & r).Value
'
'                     Sheets("volvo_NewPrices").Activate
'                     Range("F" & s).Select
'                     MsgBox ws2.Range("F" & s).Value
'
                    
                ' ws1.Range("P" & r).Value = Val(ws1.Range("L" & r)) * Val(ws2.Range("D" & s))
                  
                  
                  
                     myStringRes = Val(ws1.Range("L" & r)) * Val(ws2.Range("D" & s))
                    myRes = Round(myStringRes, 2)
                    ws1.Range("P" & r).Value = "" & myRes
                        
                 
            End If
        Next s
    Next r
End Sub


Sub ConvToNum()

'This function makes each result in a cell to format of 2 decimals;;;;;;;;;;;;;;;;;;;;;;
For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
  
  
  
  myResAB = Round((Cells(i, "AB").Value), 2)
   Cells(i, "AB").Value = myResAB
   
   
    myResN = Round((Cells(i, "N").Value), 2)
    Cells(i, "N").Value = myResN
    
    myResO = Round((Cells(i, "O").Value), 2)
    Cells(i, "O").Value = myResO
    
    myResP = Round((Cells(i, "P").Value), 2)
    Cells(i, "P").Value = myResP
   
Next i
End Sub

Sub AutoFilICCoE()

'This function makes each result in a cell to format of 2 decimals;;;;;;;;;;;;;;;;;;;;;;
For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row

   Cells(i, "Q").Value = "0"
Next i
End Sub


Sub InsertFirstRow()


    Sheets("Volvo_Row_one").Select
    Rows("1:1").Select
    Selection.Copy
    Sheets("Volvo_Statistik").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
Rows("1:1").EntireRow.Interior.Color = xlNone
End Sub

