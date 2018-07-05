Attribute VB_Name = "Module1"
Sub sheofs()


For Each sht In ActiveWorkbook.Worksheets
Set Rng = sht.UsedRange


Set MyRange = Rng
For Each MyCol In MyRange.Columns
For Each MyCell In MyCol.Cells
'MsgBox ("Address: " & MyCell.Address & Chr(10) & "Value: " & MyCell.Value)
'
'
If MyCell.Interior.ColorIndex = 23 Then



'MsgBox "Language is: " & MyCol.Cells(1, 1).Text


'MsgBox "" & Mycell.Column


Cells(MyCell.Row, 2).Copy
MyCell.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
SkipBlanks:=False, Transpose:=False
MyCell.Font.ColorIndex = 3


End If








Next
Next
Next





End Sub





