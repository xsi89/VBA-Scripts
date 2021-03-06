Attribute VB_Name = "V_Statistics"


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
    
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Sheets("Blad1").Delete
ThisWorkbook.Sheets("Orders").Delete
On Error GoTo 0
Application.DisplayAlerts = True
'     Dim sh As Worksheet
'    For Each sh In Sheets
'        If IsEmpty(sh.UsedRange) Then sh.Delete
'    Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


Sub GetColData()

  Dim X As Long
  Dim Sht As Worksheet
  Dim Target As Worksheet
  Dim VS() As String
  Dim TS() As String
  
  Const Vehicles = "Volvo_3P,Volvo_Penta,Volvo_Business_Service,Volvo_Group_Trucks_Technology,Volvo_Information_Technology_AB,Volvo_Group_Sweden,Volvo_IT"
  
  Set Target = Sheets("Volvo_Statistik")
  
  VS = Split("A,C,D,E,G,H,I,J,K,K,M,O", ",")
  TS = Split("A,D,H,I,J,K,L,M,E,F,U,AB", ",")
  
        For Each Sht In ThisWorkbook.Worksheets
            If InStr(Vehicles, "," & Sht.Name) Then
                For X = 0 To UBound(VS)
                    Sht.Columns(VS(X)).Copy Target.Cells(1, TS(X))
                Next
            End If
        Next
End Sub



Sub myyear()
'This function removes the 6 last characters from the word and then replace with the first 4 characters.Sheets("Volvo_Statistik").Activate
    For Each Sht In ThisWorkbook.Worksheets
    mysheetname = Sht.Name
    
        If Sht.Name Like "Volvo_Statistik*" Then
        
        activeWS = ActiveSheet.UsedRange.Rows.Count
        Dim cell As Range
        Dim mystring As String
            For Each cell In Range("E1:E" & activeWS)
            myResult = Left(cell.Text, Len(cell.Text) - 6)
            cell.Value = myResult
            Next
            
        Else
        MsgBox "Sheetet Volvo_Statistik Does not exists"
        End If
    Exit For
    
    
    Next Sht
    
End Sub



Sub MyMonthPone()
'This function removes

    For Each Sht In ThisWorkbook.Worksheets
    mysheetname = Sht.Name
    
        If Sht.Name Like "Volvo_Statistik*" Then
        
        activeWS = ActiveSheet.UsedRange.Rows.Count
        Dim cell As Range
        Dim mystring As String
            
            For Each cell In Range("F1:F" & activeWS)
            myResult = Left(cell.Text, Len(cell.Text) - 3)
            
           '  MsgBox myresult
            cell.Value = myResult
            Next
        Else
        MsgBox "sheets:Volvo_Statistik doest not exists"
        End If
    Exit For
    Next Sht

End Sub

Sub MyMonthPTwo()

    For Each Sht In ThisWorkbook.Worksheets
    mysheetname = Sht.Name
    
            If Sht.Name Like "Volvo_Statistik*" Then
            
            activeWS = ActiveSheet.UsedRange.Rows.Count
            Dim cell As Range
            Dim mystring As String
            
                For Each cell In Range("F1:F" & activeWS)
                myResult = Right(cell.Text, Len(cell.Text) - 5)
                'MsgBox myresult
                cell.Value = "" & myResult
                Next
            Else
            MsgBox "sheets:Volvo_Statistik doest not exists"
            End If
        Exit For
    Next Sht

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
  
    End With
    
End Sub



Sub CheckIND()



'IND=IND CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "H") = "IND" And Cells(i, "I") = "IND" And Len(Cells(i, "B")) <> 0 And Cells(i, "AB") <> 0 Then
    Cells(i, 1).EntireRow.Interior.ColorIndex = 6

    myCell = Cells(i, "B").Text
    mycellRes = myCell - 1
    DivAB = Cells(i, "AB").Text
    mytot = Round(DivAB / mycellRes)
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
    RangeToUpdate.Offset(, 17).Value = mytot
    End If
Next i
 Rows(1).EntireRow.Delete

End Sub

Sub CheckMLY()

'IND=IND CHECK FIL;;;;;;;;;;;;;;;;;;;;;;
For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    If Cells(i, "H") = "MLY" And Cells(i, "I") = "MLY" And Len(Cells(i, "B")) <> 0 And Cells(i, "AB") <> 0 Then
    Cells(i, 1).EntireRow.Interior.ColorIndex = 4

myCell = Cells(i, "B").Text
    mycellRes = myCell - 1
    DivAB = Cells(i, "AB").Text
    mytot = Round(DivAB / mycellRes)
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
    RangeToUpdate.Offset(, 17).Value = mytot
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
            
             If ws.Name = "Volvo_Group_Trucks_Technology" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo Group Trucks Technology"
            Next i
            End If
            
             If ws.Name = "Volvo_Information_Technology_AB" Then
            For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, "B") = "Volvo Information Technology AB"
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
            
                'MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                'Sheets("Volvo_Statistik").Activate
                'Range("J" & r).Select
                'MsgBox ws1.Range("J" & r).Value
                'Sheets("volvo_NewPrices").Activate
                'Range("G" & s).Select
                'MsgBox ws2.Range("G" & s).Value
                'ws1.Range("N" & r).Value = Val(ws1.Range("J" & r)) * Val(ws2.Range("G" & s))
                                                                      
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
                                 
                'MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                'Sheets("Volvo_Statistik").Activate
                'Range("J" & r).Select
                'MsgBox ws1.Range("K" & r).Value
                'Sheets("volvo_NewPrices").Activate
                'Range("F" & s).Select
                'MsgBox ws2.Range("F" & s).Value
                
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
            
                    'MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                    'Sheets("Volvo_Statistik").Activate
                    'Range("J" & r).Select
                    'MsgBox ws1.Range("K" & r).Value
                    'Sheets("volvo_NewPrices").Activate
                    'Range("F" & s).Select
                    'MsgBox ws2.Range("F" & s).Value
                    'ws1.Range("P" & r).Value = Val(ws1.Range("L" & r)) * Val(ws2.Range("D" & s))
                                          
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


Sub myAA()
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
            
                NoNWords = Val(Cells(r, "J").Value)
                FuzzyMatch = Val(Cells(r, "K").Value) ' ordantal
                HundAndRep = Val(Cells(r, "L").Value)
                
                
                NowordsS2 = Val(ws2.Range("G" & s).Value)
                FuzzyMatchS2 = Val(ws2.Range("F" & s).Value) 'resp
                HundAndRepS2 = Val(ws2.Range("E" & s).Value)
                
                
                NoWorsdsTotL = Round((NoNWords * NowordsS2), 2)
                FuzzyMatchTotL = Round((FuzzyMatch * FuzzyMatchS2), 2)
                HundAndRepTotL = Round((HundAndRep * HundAndRepS2), 2)
                
                myTot1 = NoWorsdsTotL + FuzzyMatchTotL + HundAndRepTotL
                
                
                NoWorsdsTotH = Round((NoNWords * 0.148), 2)
                FuzzyMatchTotH = Round((FuzzyMatch * 0.074), 2)
                HundAndRepTotH = Round((HundAndRep * 0.037), 2)
                
                myTot2 = NoWorsdsTotH + FuzzyMatchTotH + HundAndRepTotH
                
                myResult = myTot2 - myTot1
                ws1.Range("AA" & r).Value = myResult
   
            End If
        Next s
Next r
End Sub


Sub myColuZ()
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
            
                NoNWords = Val(Cells(r, "J").Value)
                FuzzyMatch = Val(Cells(r, "K").Value)
                HundAndRep = Val(Cells(r, "L").Value)
                myNePric = Val(ws2.Range("G" & s).Value)
                
                wordsTot = (NoNWords + FuzzyMatch + HundAndRep)
                
                myS = wordsTot * myNePric
                myAB = Cells(r, "AB").Value
                
                myRes = Round((myS - myAB), 2)
                
                ws1.Range("Z" & r).Value = myRes
   
            End If
        Next s
    Next r
End Sub

Sub ColorNegativeVal()

'This function makes each result in a cell to format of 2 decimals;;;;;;;;;;;;;;;;;;;;;;
For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
  
    If Cells(i, "Z").Value < 0 Then
    Cells(i, "Z").EntireRow.Interior.ColorIndex = 3
    End If

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

