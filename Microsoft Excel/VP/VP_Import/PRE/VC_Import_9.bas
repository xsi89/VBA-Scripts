Attribute VB_Name = "VolvoPenta_Import_6"
Sub LoopFiles()

Dim myPath As String
Dim StrCurrentfile As String
Dim StrFName As String

Dim SourceFile As String
Dim TargetFile As String

StrDocPAth = "C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org - Copy\"
StrCurrentfile = Dir(StrDocPAth & "*NoTrans.xls")
Do While StrCurrentfile <> ""


myNoTransfile = StrDocPAth & StrCurrentfile

myLangFile = Replace(StrCurrentfile, "_NoTrans", "")


MsgBox myLangFile


MsgBox myNoTransfile

StrCurrentfile = Dir

Loop

End Sub




Sub Contentpaste()

Set SourceFile = Workbooks.Open("C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org\1\UCHP_Translation2_jeeves_sv_NoTrans.xls")
Columns(1).Select
Selection.Copy

Set TargetFile = Workbooks.Open("C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org\1\UCHP_Translation2_jeeves_sv.xls")
TargetFile.Sheets.Add(After:=TargetFile.Sheets(TargetFile.Sheets.Count)).Name = "WordNotTrans"
ActiveSheet.Paste


'# hämtar celler#

ActiveWorkbook.Worksheets("Translated").Activate
Rows("1:1").Select
Selection.EntireRow.Hidden = False


ActiveWorkbook.Worksheets("WordNotTrans").Activate

Dim s As String
Dim Current As Worksheet
For Each Current In Worksheets
For Each C In ActiveSheet.UsedRange
If C.Interior.ColorIndex = 3 Then
s = C.Address
C.Copy Sheets("Translated").Range(s)
End If
Next
Next


Worksheets("WordNotTrans").Delete
Application.DisplayAlerts = True
SourceFile.Close SaveChanges:=False
End Sub





