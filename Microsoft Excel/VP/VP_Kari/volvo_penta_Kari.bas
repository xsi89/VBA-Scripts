Attribute VB_Name = "volvo_penta_Kari"
Sub CopyBtoC()

For Each sht In ActiveWorkbook.Worksheets
    For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(i, 2).EntireColumn.Select
            Selection.Copy
            For x = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            Cells(x, 3).EntireColumn.Select
            ActiveSheet.Paste
        Next x
    Next i
Next
End Sub



Sub CopyB_C()
' Copy all Active Cells from column B to column C

For i = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    Cells(i, 2).EntireColumn.Select
    Selection.Copy
    For x = 2 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
        Cells(x, 3).EntireColumn.Select
        ActiveSheet.Paste
    Next x
Next i


End Sub


Sub Splitsheets()

Dim XPath As String

myOrgName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))

XPath = Application.ActiveWorkbook.Path

MkDir XPath & "\FileSheets"
myLangPath = XPath & "\FileSheets"
MkDir XPath & "\LangCombs"


Application.ScreenUpdating = False
Application.DisplayAlerts = False
    For Each xWs In ThisWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs fileName:=myLangPath & "\" & myOrgName & "_" & xWs.Name & ".xls"
    Application.ActiveWorkbook.Close False
    Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

