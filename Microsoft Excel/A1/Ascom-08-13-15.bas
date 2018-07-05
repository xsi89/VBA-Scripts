Attribute VB_Name = "Ascom"
Sub myCOl()
'written by XsiSec 2015-08-15
Dim i As Long, j As Long, arrCol()
arrCol = Array("K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V")
For i = 2 To Range("D" & Rows.Count).End(xlUp).Row
    For j = 0 To UBound(arrCol)
        If Cells(i, "D").Value = "NONE" Then
            Cells(i, arrCol(j)).Interior.ColorIndex = 4
                Else
            If IsNumeric(Cells(i, "D").Value) Then
                If CLng(Cells(i, "D").Value) > Len(Cells(i, arrCol(j)).Value) Then
                    Cells(i, arrCol(j)).Interior.ColorIndex = 4
                        Else
                    Cells(i, arrCol(j)).Interior.ColorIndex = 3
                                        
                End If
            End If
        End If
    Next j
Next i
End Sub

