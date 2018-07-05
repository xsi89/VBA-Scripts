Attribute VB_Name = "MT_Feedback_Report"
Sub myrun()

Form.Show

End Sub



Sub ConvNum()

For i = 3 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
  
        If Cells(i, "H") > "" Then
        myResN = Round((Cells(i, "H").Value), 2)
        Cells(i, "H").Value = myResN
        End If
        
        

Next i



End Sub


Sub DelLasRow()

Dim lRow As Long
Dim lCol As Long
    
    lRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    Rows(lRow).ClearContents
    
    
End Sub
