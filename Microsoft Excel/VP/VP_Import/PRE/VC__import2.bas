Attribute VB_Name = "Volvo_Penta_import"

Sub Move_Colored_Cells()
Dim s As String
For Each c In ActiveSheet.UsedRange
If c.Interior.ColorIndex = 3 Then
s = c.Address
c.Copy Sheets("Sheet1").Range(s)
End If
Next
End Sub


