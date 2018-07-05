Attribute VB_Name = "Find_Lang"
Sub Find_Language()
For Each sht In ActiveWorkbook.Worksheets

'MsgBox sht.Name






    Dim lastRow As Long, LastCol As Integer, c As Integer, r As Long
     
    lastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    LastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Column
     
    For c = 1 To LastCol
        For r = 1 To lastRow
            If Cells(r, c).Interior.ColorIndex = 23 Then
            
                 
                 
                 MsgBox "the cell row is " & Cells(c).Text
                 
                 ' Your code
            End If
        Next r
    Next c

Next




End Sub


'säger blå cell innehåll




Sub test()

For Each sht In ActiveWorkbook.Worksheets
Set Rng = sht.UsedRange




For Each rw In Rng.Rows
For Each cl In rw.Columns

If cl.Interior.ColorIndex = 23 Then

MsgBox cl.Text

End If
Next cl
Next rw


Next sht

End Sub











Sub Langauge_Combination()

mypath = ActiveWorkbook.Path

For Each sht In ActiveWorkbook.Worksheets
    Set Rng = sht.UsedRange
    Set MyRange = Rng

        For Each MyCol In MyRange.Columns
        For Each myCell In MyCol.Cells
        
            If myCell.Interior.ColorIndex = 23 Then
            sht.Cells(myCell.Row, MyRange.Columns(2).Column).Copy


                'MsgBox "Language is: " & MyCol.Cells(1, 1).Text
                Dim objSheet As Worksheet
                For Each objSheet In ActiveWorkbook.Sheets
                   If objSheet.Name = MyCol.Cells(1, 1).Text Then
                        End 'MsgBox "A worksheet with that name already exists."
                Worksheets(MyCol.Cells(1, 1).Text).Activate
                    End If
                      Next objSheet
                      
                      
    Set NewBook = Workbooks.Add
    myName = MyCol.Cells(1, 1).Text
    
    


ActiveWorkbook.SaveAs Filename:=(mypath) & "tet.xls"

   ' With NewBook
    
   '   .Title = "All Sales"
    '    .Subject = "Sales"
  '     .SaveAs fileName:=FilePath & "\" & fileName ' (without the ' ) in your ActiveWorkbook.Save?
 '  End With
                      
                            
                   

            End If
            

        Next
    Next
Next
End Sub



Sub Copier4()
   Dim x As Integer

   For x = 1 To ActiveWorkbook.Sheets.Count
      'Loop through each of the sheets in the workbook
      'by using x as the sheet index number.
      ActiveWorkbook.Sheets(x).Copy _
         After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
         'Puts all copies after the last existing sheet.
   Next
End Sub



Sub CopyWorkbook()

Dim sh As Worksheet, wb As Workbook

Set wb = Workbooks("Target workbook")
For Each sh In Workbooks("source workbook").Worksheets
   sh.Copy After:=wb.Sheets(wb.Sheets.Count)
Next sh

End Sub



Sub Volvopentatest()


'mypath = "C:\"
MFileN = ActiveWorkbook.Name
myfilename = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
   


For Each sht In ActiveWorkbook.Worksheets
    Set Rng = sht.UsedRange
    Set MyRange = Rng


        For Each MyCol In MyRange.Columns
        For Each myCell In MyCol.Cells
        
       
            If myCell.Interior.ColorIndex = 23 Then
         
            myName = MyCol.Cells(1, 1).Text




                      


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


                     



  Dim wb As Workbook, x As String
   For Each wb In Workbooks
    '  If wb.Name <> ThisWorkbook.Name Then
         x = wb.Name
         Workbooks(x).Activate
       
       
       
       NewWBName = ActiveWorkbook.Name
       
       
       

       
       MsgBox NewWBName
       
       
       
    '   MyAname = Split(NewWBName, ".xls")
       
       
       
       
       
       

'        If myCell.Interior.ColorIndex = 23 Then
'
'            MsgBox myName = MyCol.Cells(1, 1).Text
'
'            End If
       
       
       
       
  '    End If
 
 
 
   Next wb






   'ActiveWorkbook.SaveAs Filename:=(mypath) & (myfilename) & "_" & (myName) & ".xls"
        


    
'& (MyCol.Cells(1, 1))= sv




'MsgBox (myfilename)


          
          


            End If
            


        Next
        
        
    Next
  
    
Next
End Sub


