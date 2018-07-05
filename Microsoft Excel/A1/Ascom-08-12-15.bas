Attribute VB_Name = "Ascom"
Sub MyCelCounter()

'written by XsiSec 08 12 15

Dim i As Long, c As Range, ArrCol() As Variant
ArrCol = Array("T", "U", "V")

    For i = LBound(ArrCol) To UBound(ArrCol)
        For Each c In Range(ArrCol(i) & "2:" & ArrCol(i) & Range(ArrCol(i) & Rows.Count).End(xlUp).Row)
        Dim myActC As Variant
        
        CharCount = Len(c.Value)
        myActC = Range("D" & c.Row).Value
        
            If (myActC > CharCount) Then
            c.Interior.ColorIndex = 4
            Else
                If (myActC < CharCount) Then
                c.Interior.ColorIndex = 3
                End If
            End If
            
                    If (Range("D" & c.Row).Value) <> "" Then
                        If myAct = "NONE" Then
                        c.Interior.ColorIndex = 4
                        End If
                    End If
                
        Next c
    Next i
    
End Sub




