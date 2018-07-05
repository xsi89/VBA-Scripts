Attribute VB_Name = "Module1"
Sub sheofs()


For Each sht In ActiveWorkbook.Worksheets
Set Rng = sht.UsedRange
Set MyRange = Rng

'Rows(1).Interior.Color = vbBlue
Rows(1).Font.Color = vbRed



For Each MyCol In MyRange.Columns
For Each myCell In MyCol.cells


If myCell.Interior.ColorIndex = 23 Then
myCell.Font.ColorIndex = 3
cells(myCell.Row, 2).Copy
myCell.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False



  
    


Next

Next
Next



 
   
   
End Sub







Sub colors()


For Each sht In ActiveWorkbook.Worksheets
Set Rng = sht.UsedRange
Set MyRange = Rng

For Each MyCol In MyRange.Columns
For Each myCell In MyCol.cells

        If myCell.Interior.ColorIndex = xlNone Then myCell.ClearContents
        
        Next
        Next
        Next
        
   
End Sub

