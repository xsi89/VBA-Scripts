Attribute VB_Name = "Volvo_Penta_Import24"
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
Dim OpenFile As String
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
               
                
               OpenFile = (myPath & myfile)

'MsgBox openfile
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





Sub test()


If Dir(myPath & "UCHP_Translation2_jeeves_fra.xls") <> "" Then
MsgBox "File exists"
Else
MsgBox "File does not exist"
End If

End Sub






























Sub ChangePR()
Application.DisplayAlerts = False

Dim wb As Workbook
Dim myPath As String
Dim myfile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
'  myExtension = "*_Tur.xls"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myfile)
    
    'gör något




Rows("1:1").Select
Selection.EntireRow.Hidden = False
myLangCol = ActiveSheet.Cells("1,1").Text

If myLangCol = "pt_BR" Then

Application.Range("A1").Value = "br"



End If


    
    'Save and Close Workbook
      wb.Close SaveChanges:=True

    'Get next file name
      myfile = Dir
  Loop

'Message Box when tasks are completed
 ' MsgBox "Nu är alla celler på Rad 1 Dolda!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub















