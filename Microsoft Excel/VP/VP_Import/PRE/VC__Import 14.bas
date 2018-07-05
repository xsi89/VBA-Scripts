Attribute VB_Name = "Module1"

Sub MergeNoTransAndLang()

Dim myPath As String
Dim StrCurrentfile As String
Dim StrFName As String
Dim myLangFile As String
Dim intResult As Integer

Application.DisplayAlerts = True


intResult = Application.FileDialog(msoFileDialogFolderPicker).Show

If intResult = 0 Then

    MsgBox "User pressed cancel macro will stop!"

Exit Sub

Else

strDocPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"

End If


StrCurrentfile = Dir(strDocPath & "*NoTrans.xls")
Do While StrCurrentfile <> ""


myNoTransfile = strDocPath & StrCurrentfile

myLangFile = Replace(StrCurrentfile, "_NoTrans", "")

'MsgBox myLangFile

Set myLangFileN = Workbooks.Open(strDocPath & StrCurrentfile)
Columns(1).Select
Selection.Copy

Set noLangFilen = Workbooks.Open(strDocPath & myLangFile)
noLangFilen.Sheets.Add(After:=noLangFilen.Sheets(noLangFilen.Sheets.Count)).Name = "WordNotTrans"
ActiveSheet.Paste



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
Application.DisplayAlerts = False
myLangFileN.Close SaveChanges:=False
noLangFilen.CheckCompatibility = False
noLangFilen.Close SaveChanges:=True

  'FileCopy "C:\Users\Ron\SourceFolder\Test.xls", "C:\Users\Ron\DestFolder\Test.xls"

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


Sub Langauge_Combination()

Dim myPath As String
Dim myLangCol As String

myPath = Application.ActiveWorkbook.Path


For Each sht In ActiveWorkbook.Worksheets
Set rng = sht.UsedRange
Set MyRange = rng


For Each MyCol In MyRange.Columns
For Each myCell In MyCol.Cells








'MsgBox ("Address: " & MyCell.Address & Chr(10) & "Value: " & MyCell.Value)


If myCell.Interior.Color = vbBlue Then

If myCell.Font.Color = vbRed Then




myLangCol = MyCol.Cells(1, 1).Text

MyFile = myPath & myLangCol


'myLangColTxtFilter = MyCol.Cells(1, 1).Text

'MsgBox MyFile


Dim xPath As String
xPath = Application.ActiveWorkbook.Path
MyFile = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
Sheetname = ActiveSheet.Name




MsgBox xPath & MyFile & "_" & Sheetname & "_"


End If

'MsgBox "" & Mycell.Column
'Cells(Mycell.Row, 2).Copy
'Mycell.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
'SkipBlanks:=False, Transpose:=False
    
            End If
        Next
    Next
Next




End Sub

