Attribute VB_Name = "Volvo_Penta_Import"
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

Sub OpenLangFile()
Dim myPath As String
Dim myfile As String
Dim mysheet As String
Dim mySpa As String
Dim myLangC As String
Dim OrgWB As Workbook
Dim openfile As String
Set OrgWB = ActiveWorkbook


myPath = Application.ActiveWorkbook.Path & "\"
mySpa = "_"
myFileN = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
SheetN = ActiveSheet.Name
myExt = ".xls"

Set wsSrc = ActiveSheet
Set rng = wsSrc.UsedRange



    For Each cl In rng.Columns
        For Each r In cl.Rows(1).Cells
            If r.Value <> "" Then

                myLangC = r.Text
                r.Select
                ActiveCell.EntireColumn.Select
                OrgCol = ActiveCell.Column
                myOrgfile = ActiveWorkbook.Name
                myfile = myFileN & mySpa & SheetN & mySpa & myLangC & myExt
                Application.DisplayAlerts = False
                ActiveWorkbook.Save
               
                
               openfile = (myPath & myfile)


'MsgBox openfile
'
                If Dir(openfile) <> "" Then
                    Set getWb = Workbooks.Open(openfile)
                    Application.DisplayAlerts = False
                    Columns(1).Select
                    Selection.Copy
                    OrgWB.Activate
                    ActiveSheet.Paste
                    getWb.Close

                    Else
                    MsgBox "File does not exist"
                End If


            End If
        Next r
    Next cl


End Sub





Sub test()


If Dir(myPath & "UCHP_Translation2_jeeves_fra.xls") <> "" Then
MsgBox "File exists"
Else
MsgBox "File does not exist"
End If

End Sub
