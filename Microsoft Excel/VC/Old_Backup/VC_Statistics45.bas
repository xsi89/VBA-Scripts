Attribute VB_Name = "Volvo_Statistics"

Sub CombineFiles()

'/=======================================================================================================
'/ When you run this function, select following folder from the structure you made "\Root\Resources\".
'/ The function select all files in the folder then combine them to current workbook also gives each
'/ imported file the filename as sheet name except the extension so example "test.xls" result "test".
'/ The function deletes also the sheet named: "Orders".
'/=======================================================================================================

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

Sub all()
'/=======================================================================================================
'/ It is a batch function to run possible functions except the combine function the first one.
'/=======================================================================================================

Call GetColData
Call mYear
Call myMonth
Call CheckInstandCol
Call CheckNumOfIns
Call RemColCells
Call Company
Call NewWordsC
Call FuzzyC
Call RepsandHund
Call ConvToNum
Call AutoFilICCoE
Call myZ
Call myAA
Call ColorNegativeVal
Call InsertFirstRow

End Sub

Sub GetColData()
'/=======================================================================================================
'/ This function it is an array that checks if one of the following sheets exists:
'/ "Volvo_3P"," Volvo_Bus"," Volvo_Business_Service","Volvo_Group_Sweden","Volvo_IT.
'/ If the sheet exists then copy certain columns from the sheet into the sheet: "Volvo_Statistik".
'/ (This is the main reason it is very important to have exactly the following names on the files).
'/=======================================================================================================

  Dim x As Long
  Dim Sht As Worksheet
  Dim Target As Worksheet
  Dim VS() As String
  Dim TS() As String
  'I use Const here to be able to have one string that contains many elements, and each element is seperated by a comma.
  Const Vehicles = "Volvo_3P,Volvo_Penta,Volvo_Business_Service,Volvo_Group_Trucks_Technology,Volvo_Information_Technology_AB,Volvo_Group_Sweden,Volvo_IT"
  
  Set Target = Sheets("Volvo_Statistik")
  
  'here is the columns from the company sheet
  VS = Split("A,C,D,E,G,H,I,J,K,K,M", ",")
  
  'here is the positioning for all columns
  TS = Split("A,D,H,I,J,K,L,M,E,F,U", ",")
  
        For Each Sht In ThisWorkbook.Worksheets
            If InStr(Vehicles, "," & Sht.Name) Then
                For x = 0 To UBound(VS)
                    Sht.Columns(VS(x)).Copy Target.Cells(1, TS(x))
                Next
            End If
        Next

End Sub

Sub mYear()

'/=======================================================================================================
'/ The code loops all the worksheets, if the sheet name "Volvo_Statistik" exists, then
'/ loop all cells in entire column "E". For each cell strip the cell value from left with 6 characters.
'/ Example 2015-03-31 equals to 2015.
'/=======================================================================================================

        For Each Sht In ThisWorkbook.Worksheets
        mysheetname = Sht.Name
        
            If Sht.Name Like "Volvo_Statistik*" Then
            
            activeWS = ActiveSheet.UsedRange.Rows.Count
            Dim cell As Range
            Dim mystring As String
                For Each cell In Range("E1:E" & activeWS)
                myresult = Left(cell.Text, Len(cell.Text) - 6)
                cell.Value = myresult
                Next
                
            Else
            MsgBox "Sheetet Volvo_Statistik Does not exists"
            End If
        Exit For

    Next Sht
    
End Sub

Sub myMonth()

'/=======================================================================================================
'/ This code change the format from 2015-02-21 to 02
'/=======================================================================================================

Dim Sht As Worksheet
    For Each Sht In ThisWorkbook.Worksheets

        If Sht.Name Like "Volvo_Statistik*" Then
            With Sht.Range("F2:F" & Sht.Range("F" & Sht.Rows.Count).End(xlUp).row)
            .NumberFormat = "@"
            .Cells = Evaluate("=INDEX(TEXT(MONTH('" & Sht.Name & "'!" & .Address & "+0),""00""),0)")
            End With
        End If
    
                For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
                    myF = Cells(i, "F").Value
                    Cells(i, "F").Value = Val(myF)
                Next i
    Next Sht
        Rows(1).EntireRow.Delete
End Sub

Sub CheckInstandCol()

'/=======================================================================================================
'/ This function contains two loops the first loop Z That checks couple of statements:
'/ If each row cell column "H" have value "MLY" then check if column "I" have value "MLY" then
'/ create a string of the cell in same row but column "A". Color then the entire row green, also present
'/ the number of instances in the same row but the cell in column "B".
'/ In the next loop "X", it is almost the same but check instead for the value "IND"
'/ then color the entire row yellow.
'/=======================================================================================================

  For z = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
  'if Value in column H and I equal "MLY" then look in whole Column A for instances
  
        If Cells(z, "H").Value = "MLY" And Cells(z, "I").Value = "MLY" Then ' statements
        mycell = Cells(z, "A").Value
        myOcc = Application.CountIf(Range("A1", Cells(Rows.Count, "A").End(xlUp)), mycell)
        myNumofOC = myOcc - 1 'Here is the number of instances less -1
        Cells(z, "B").Value = myNumofOC ' set the number of how many instances on the same row
        
        Cells(z, "B").EntireRow.Interior.ColorIndex = 4
        End If
        
            Next z
            'END LOOP
            
            'Below it's a loop(this thime without any statements) and then use the order number (cells(Z,"B").value)
            'then look for it everywhere in same column
                For x = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
                
                    If Cells(x, "H").Value = "IND" And Cells(x, "I").Value = "IND" Then
                    
                    mycell = Cells(x, "A").Value
                    myOcc = Application.CountIf(Range("A1", Cells(Rows.Count, "A").End(xlUp)), mycell)
                    myNumofOC = myOcc - 1
                    Cells(x, "B").Value = myNumofOC
                    myOrderNr = Cells(x, "A").Value
                    myPrelCost = Cells(x, "AB").Value
                    myresult = myPrelCost / myNumofOC
                    Cells(x, "B").Value = myNumofOC
                    Cells(x, "B").EntireRow.Interior.ColorIndex = 6
                    
                    End If
                Next x

End Sub

Sub CheckNumOfIns()

'/=======================================================================================================
'/ the first loop "I" checks if the value in every cell in each row column "B" have a greater value than "0"
'/ then create a couple of strings on the same row but from different columns.
'/ The basic idea behind the math formula it is to use the value from the column "B"
'/(the cell presents the number of instances) divide then the value by the cell in same row in column "AB".
'/ There is then a new loop that checks each cell in column "A" every row if the text value is equal to "0"
'/ paste then math formula summarize in column "R".
'/=======================================================================================================

    For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    
        If Cells(i, "b").Value > 0 Then
        mycell = Cells(i, "A").Value
        myVal = Cells(i, "B").Value
        myPrel = Cells(i, "AB").Value
        mysum = Val(myPrel / myVal)
        End If
      
            For x = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
            
                If Cells(x, "A").Value = mycell Then
                Cells(x, "R").Value = mysum
                End If
                
            Next x
    
    Next i
    
End Sub

Sub RemIND()

'/=======================================================================================================
'/ The code loops all cells in column "H" if the text content is "IND" then delete the entire row.
'/=======================================================================================================
    For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
        If Cells(i, "H") = "IND" Then
        Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
      
End Sub

Sub RemMLY()

'/=======================================================================================================
'/ The code loops all cells in column "H" if the text content is "MLY" then delete the entire row.
'/=======================================================================================================


  For x = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
            If Cells(x, "H") = "MLY" Then
            Cells(x, 1).EntireRow.Delete
            End If
        Next x

End Sub


Sub myCompany()

'/=======================================================================================================
'/ The code clear the entire column "B".
'/ For each sheet in the workbook check for following names:
'/ Volvo_3P , Volvo_Penta, Volvo_Bus, Volvo_Business_Service, Volvo_Group, Sweden, Volvo_IT
'/ when one of the names are found use the value auto fill all used cells in entire column "B".
'/=======================================================================================================

Dim ws As Worksheet
Dim myStrArray As Variant

    For Each ws In ActiveWorkbook.Worksheets
    
        myStrArray = Array("Volvo_3P", "Volvo_Penta", "Volvo_Bus", "Volvo_Group_Trucks_Technology", "Volvo_Information_Technology_AB", "Volvo_Business_Service", "Volvo_Group_Sweden", "Volvo_IT")
        
            For i = LBound(myStrArray) To UBound(myStrArray)
            mycol = myStrArray(i)
            myWsN = ws.Name
            
                If ws.Name = mycol Then
                Worksheets("Volvo_Statistik").Activate
                
                    For x = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
                        myresult = Replace(mycol, "_", " ")
                        Cells(x, "B").Value = myresult
                    Next x
                    
                End If
                
            Next i
        Next ws
End Sub


Sub NewWordsC()

'/=======================================================================================================
'/ this code compare couple statements in two different sheets.
'/ First loop all cells in column "H" and "I" in sheet "Volvo_Statistik" as long the two text cell values
'/ same row are found in the sheet "Volvo_NewPrices" column "A" and "B" same row but the cells can be
'/ found anywhere in the sheet as long they are on same row. Then go back to the sheet "Volvo_Statistik"
'/ column "J" multiply the value with the found cells value in sheet "Volvo_NewPrices" but column "G"
'/ set the result in sheet "Volvo_Statistik" same row but column "N".
'/=======================================================================================================

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim N As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    N = ws2.Range("A" & ws2.Rows.Count).End(xlUp).row
    
    For r = 1 To m
        For s = 1 To N
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

'/=======================================================================================================
'/ This code is very similar to the function "NewWordsC"
'/ First loop all cells in column "H" and "I" in sheet "Volvo_Statistik"
'/ as long the two text cell values same row are found in the sheet "Volvo_NewPrices" column "A" and "B"
'/ same row but the cells can be found anywhere in the sheet as long they are on same row.
'/ Then go back to the sheet "Volvo_Statistik" column "F" multiply the value with the found cells value
'/ in sheet "Volvo_NewPrices" but column "K" set the result in sheet "Volvo_Statistik" same row but column "O".
'/=======================================================================================================
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim N As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    N = ws2.Range("A" & ws2.Rows.Count).End(xlUp).row
    
    For r = 1 To m
        For s = 1 To N
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

'/=======================================================================================================
'/ This code is very similar to the function "NewWordsC"
'/ First loop all cells in column "H" and "I" in sheet "Volvo_Statistik" as long the two text cell values
'/ same row are found in the sheet "Volvo_NewPrices" column "A" and "B" same row but the cells can be
'/ found anywhere in the sheet as long they are on same row. Then go back to the sheet "Volvo_Statistik"
'/ Column "L" multiply the value with the found cells value in sheet "Volvo_NewPrices" but column "D"
'/ set the result in sheet "Volvo_Statistik" same row but column "P".
'/=======================================================================================================

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim N As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    N = ws2.Range("A" & ws2.Rows.Count).End(xlUp).row
    
    For r = 1 To m
        For s = 1 To N
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

Sub CountU()
'/=======================================================================================================
'/ The loop "I" if the cell each row column "R" have a greater value than 0 then create variable "myCell"
'/ also create a variable "myMulti" equals to 42, and "myRes"= myCell * myMulti. Present the new value in
'/ same cell "column "U"
'/=======================================================================================================


    For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
        If Cells(i, "U").Value > 0 Then
            mycell = Cells(i, "U").Value
            myMulti = 42
            myRes = mycell * myMulti
            Cells(i, "U").Value = myRes
        End If
    Next i

End Sub




Sub myColuZ()

'/=======================================================================================================
'/ This function contains two loops ("R" and "S").
'/ When you run function do the function first activates the current sheet ("Volvo_Statistik"), give label
'/ "ws1". Then the function call the loop "R" that loops all certain cells for each row (columns "H" and "I").
'/ Activate then "Volvo_NewPrices" give label "ws2", then runs the loop called "S" check in the columns
'/ "A" and "B" and if these two cells from loop "R" matches the cells content in columns "A" and "B" "ws2"
'/ same row. (see example below)
'/
'/ "Volvo_Statistik" row 1: column H ="Volvo_NewPrices" row 1: column A content
'/ "Volvo_Statistik" row 1: column I="Volvo_NewPrices" row 1: column B contents
'/
'/ next section of the code is hard to explain, but basically the function creates a couple of variables
'/ that the function multiple differently (easier to see in the code) then present the new value in column "Z"
'/=======================================================================================================


 Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim N As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    N = ws2.Range("A" & ws2.Rows.Count).End(xlUp).row
    
    For r = 1 To m
        For s = 1 To N
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
                
                NoWorsdsTotH = Round((NoNWords * NowordsS2), 2)
                FuzzyMatchTotH = Round((FuzzyMatch * NowordsS2), 2)
                HundAndRepTotH = Round((HundAndRep * NowordsS2), 2)
                
                myTot2 = NoWorsdsTotH + FuzzyMatchTotH + HundAndRepTotH
                
                myresult = myTot2 - myTot1
                
                
                ws1.Range("Z" & r).Value = myresult
                
              '  Debug.Print myTot2 - myTot1
   
            End If
        Next s
Next r
End Sub


Sub myColuAA()

'/=======================================================================================================
'/This function take all the values from 3 different columns and multiply with the old prices in "Kr" then
'/ the total cost and put the result in column AA.
'/=======================================================================================================

 Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim N As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    N = ws2.Range("A" & ws2.Rows.Count).End(xlUp).row
    
    For r = 1 To m
        For s = 1 To N
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
                
                myresult = myTot2 - myTot1
                ws1.Range("AA" & r).Value = myresult
   
            End If
        Next s
Next r
End Sub



Sub myContMts()
'/=======================================================================================================
'/ The basic idea of the loops in this function is to recreate the analyze filename by using certain
'/ values from the different columns in the same row. I also declare a few strings that are not listed
'/ in the columns. The next loop uses the string that is generated from the earlier loop to open the
'/ analyze file in location C:\data\. The next loop checks for the text "/batchTotal/analyse/exact/@words/#agg".
'/ When text is found loop the active column and if the value is equal to "0" or greater than "0" copy the
'/ value to the "Volvo_Statistik" workbook same row but column "M".
'/=======================================================================================================
Dim myName As String
Dim MyFolder As String
Dim myfile As String
Dim wb As Workbook
Set wb = ActiveWorkbook
Dim Wbt As Workbook
Dim myContM As String
Dim myContMRes As String
Application.EnableCancelKey = xlDisabled

    For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    MyFolder = "C:\DATA\"
    mycell = Cells(i, "A").Value
    myAnalyzF = "_Analyze_Files_"
    mySoLang = Cells(i, "H").Value
    myHyph = "-"
    myTarLang = Cells(i, "I").Value
    myExt = ".xml"
    myName = mycell & myAnalyzF & mySoLang & myHyph & myTarLang & myExt
    myfile = MyFolder & myName
    
    If Dir(myfile) <> "" Then
        Set Wbt = Workbooks.Open(myfile)
        
            For Each N In ActiveSheet.UsedRange
                If N.Value = "/batchTotal/analyse/inContextExact/@words/#agg" Then
                N.Select
                
                    For x = 3 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
                    
                        If Cells(x, ActiveCell.Column).Value = "0" Then
                            myContextM = Cells(x, ActiveCell.Column).Copy
                            Else
                             
                             If Cells(x, ActiveCell.Column).Value > 0 Then
                             
                             myContextM = Cells(x, ActiveCell.Column).Copy
                            
                            'MsgBox myContextM = Cells(x, ActiveCell.Column).Value
                            
                        End If
                        
                        End If
                                      
                    Next x
                End If
                
            Next N
            
        wb.Activate
        Cells(i, "M").Activate
        ActiveSheet.Paste
        Wbt.Close
        Else
             
       'Debug.Print myfile & " does not exist."
            
        End If
    Next i
    
    Application.EnableCancelKey = xlenabled

End Sub



Sub myCMDiff()

'/=======================================================================================================
'/ This function is constructed very similar to the function "myContextMts
'/ First loop all cells in column "H" and "I" in sheet "Volvo_Statistik" as long the two text cell values
'/ same row are found in the sheet "Volvo_NewPrices" column "A" and "B" same row but the cells can be
'/ found anywhere in the sheet as long they are on same row. Then multiple the cell ws1 column "M" with
'/ the cell value in ws2 column D and place new the value in ws1 column "Q".
'/=======================================================================================================
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim r As Long
    Dim m As Long
    Dim s As Long
    Dim N As Long

    Set ws1 = Worksheets("Volvo_Statistik")
    m = ws1.Range("A" & ws1.Rows.Count).End(xlUp).row
    
    Set ws2 = Worksheets("volvo_NewPrices")
    N = ws2.Range("A" & ws2.Rows.Count).End(xlUp).row
    
    For r = 1 To m
        For s = 1 To N
            If Trim(ws1.Range("H" & r) & ws1.Range("I" & r)) = Trim(ws2.Range("A" & s) & ws2.Range("B" & s)) Then
            
                    'MsgBox "Cells " & "A" & r & " " & "B" & r & " on Sheet1 contain " & ws1.Range("A" & r) & " " & ws1.Range("B" & r) & " which matches A" & s & " " & "B" & s & " on Sheet2"
                    myStringRes = Val(ws1.Range("M" & r)) * Val(ws2.Range("D" & s))
                    myRes = Round(myStringRes, 2)
                    ws1.Range("Q" & r).Value = "" & myRes
                        
                 
            End If
        Next s
    Next r
End Sub


Sub ConvToNum()
'/=======================================================================================================
'/ This function converts following values from certain columns "N","O","P" and "Q" to regular numbers
'/=======================================================================================================

For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
  
        myResN = Round((Cells(i, "N").Value), 2)
        Cells(i, "N").Value = myResN

        myResO = Round((Cells(i, "O").Value), 2)
        Cells(i, "O").Value = myResO

        myResP = Round((Cells(i, "P").Value), 2)
        Cells(i, "P").Value = myResP
        
    myResP = Round((Cells(i, "Q").Value), 2)
        Cells(i, "Q").Value = myResP
Next i
End Sub


Sub noAnFiles()
'/=======================================================================================================
'/ This code loop every row, if the text-content in column "Q" is "0" start loop in column "M" then if
'/ the cell is empty color cell red.
'/=======================================================================================================

    For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
        If Cells(i, "Q").Value = "0" Then
            If Cells(i, "M").Value = "" Then
                Cells(i, "M").EntireRow.Interior.ColorIndex = 3
            End If
        End If
    Next i

End Sub

Sub myColAB()
'/=======================================================================================================
'/ Each letter in the array (N, O, P, Q, R, S, T, U, V, W, X, Y).presents a column, the loop loops
'/ through all the columns each row and for each cell in all certain columns summarize the cells to
'/ a total and present the new value in column "AB" same row.
'/=======================================================================================================

Dim x As Long
Dim y As Double
Dim i As Long
Dim j As Long
Dim myArr

x = Cells(Rows.Count, "N").End(xlUp).row

myArr = Array("N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y")

    For i = 1 To x
    z = 0
        For j = LBound(myArr) To UBound(myArr)
                z = z + Range(myArr(j) & i).Value
        Next j
            Range("AB" & i) = z
    Next i

End Sub

Sub ColorNegativeVal()

'/=======================================================================================================
'/ This code loop every cell in column "Z" each row and if the cell value is less than 0 then color it 3
'/=======================================================================================================

For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
  
    If Cells(i, "Z").Value < 0 Then
    Cells(i, "Z").EntireRow.Interior.ColorIndex = 3
    End If
    
Next i
End Sub

Sub InsertFirstRow()
'/=======================================================================================================
'/ This code selects the first row in the sheet "Volvo_Row_One", then copying
'/ select sheet "Volvo_Statistik" paste first, also removes color of the row.
'/=======================================================================================================
    
    Sheets("Volvo_Row_one").Select
    Rows("1:1").Select
    Selection.Copy
    Sheets("Volvo_Statistik").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
Rows("1:1").EntireRow.Interior.Color = xlNone

End Sub

Sub FilEmptyCells()

'/=======================================================================================================
'/ This is an array that contain specific columns ("Q", "R", "S","T", "V", "W", "X", "Y") the Loop "X"
'/ for each empty cells.
'/=======================================================================================================

Dim myStrArray As Variant

myStrArray = Array("Q", "R", "S", "T", "V", "W", "X", "Y")
    
    For i = LBound(myStrArray) To UBound(myStrArray)
    
    mycol = myStrArray(i)
        For x = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
            
            If Cells(x, mycol).Value = "" Then
                
               ' Cells(x, mycol).Select
               ' MsgBox Cells(x, mycol).Value
                Cells(x, mycol).Value = "0"
            
            End If
        
        Next x
    
    Next i

End Sub
