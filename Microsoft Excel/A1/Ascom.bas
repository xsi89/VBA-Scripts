Attribute VB_Name = "Ascom"
Sub AscomCountColumns()
'Written by XsiSec 2015-08-11

        
 Dim i As Long
 
 For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
 
 

 columnD = Cells(i, 4).Value
 
 'MsgBox columnD
 
 

If (columnD = Cells(i, 4).Value) <> "" Then

'MsgBox columnD


 EngCol = Len(Cells.Item(i, "V").Text)

If columnD > EngCol Then


mystringtwo = Cells.Item(i, "V").Value

Cells.Item(i, "V").Interior.ColorIndex = 4

'MsgBox "mindre"

Else
Cells.Item(i, "V").Interior.ColorIndex = 3
End If

End If
Next i
    
End Sub


