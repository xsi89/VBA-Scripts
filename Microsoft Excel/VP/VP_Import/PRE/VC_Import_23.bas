Attribute VB_Name = "Volvo_PentaImport_22"
Sub MergeNoTransAndLang()
'sätter ihop Notrans i _sv filer
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


StrCurrentfile = Dir


Loop


End Sub

Sub MergeColumns()
Dim myPath As String
Dim mysheet As String
Dim mySpa As String
Dim myLangC As String
Dim OrgWB As Workbook
Set OrgWB = ActiveWorkbook
Dim myFile As String
Dim OpenFile As String

' sätter ihop columner på rätt ställe
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
                myOrgFile = ActiveWorkbook.Name
                
                
                myFile = myFileN & mySpa & myLangC & myExt
                Application.DisplayAlerts = False
                ActiveWorkbook.Save
                OpenFile = myPath & myFile
                
            '    MsgBox OpenFile
'
                If Dir(OpenFile) <> "" Then
                    Set getWb = Workbooks.Open(OpenFile)
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








Sub myfileopener()

Dim File As String
Dim StrFName As String
Dim myPath As String
Dim intResult As Integer

Application.DisplayAlerts = True
intResult = Application.FileDialog(msoFileDialogFolderPicker).Show
If intResult = 0 Then
MsgBox "Du avbröt macrot"
Exit Sub

Else
myPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) & "\"

End If


File = Dir(myPath & "*.xls")
Do While File <> ""


If Not File Like "*_??.*" Then


Call MergeColumns






          '  Workbooks.Open (File)
            ' Put code to process files here.
            
            
             File.Close SaveChanges:=True
            'MsgBox File
            
        End If

File = Dir


Loop

End Sub

