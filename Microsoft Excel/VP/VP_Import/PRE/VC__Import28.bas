Attribute VB_Name = "Volvo_Penta_Import27"
Sub Change_pt_OrgFile()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Set wsSrc = ActiveSheet
Set rng = wsSrc.UsedRange

  For Each cl In rng.Columns
        For Each r In cl.Rows(1).Cells
            If r.Value = "pt_BR" Then

                r.Value = "br"
            End If
        Next r
    Next cl
Next ws
End Sub

Sub Remove_pr()
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

      wb.Close SaveChanges:=True
    'Get next file name
      myfile = Dir
  Loop

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub





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
Dim myfile As String
Dim OpenFile As String
Dim fldr As FileDialog


  Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then
            Exit Sub
        End If
        myPath = .SelectedItems(1) & "\"
    End With

' sätter ihop columner på rätt ställe
'myPath = Application.ActiveWorkbook.Path & "\"
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
                myfile = myFileN & mySpa & SheetN & mySpa & myLangC & myExt
                Application.DisplayAlerts = False
                ActiveWorkbook.Save
                OpenFile = myPath & myfile
                
                
            
                
               MsgBox OpenFile

'                If Dir(OpenFile) <> "" Then
'                    Set getWb = Workbooks.Open(OpenFile)
'                    Application.DisplayAlerts = False
'                    Columns(1).Select
'                    Selection.Copy
'                    OrgWB.Activate
'                    ActiveSheet.Paste
'                    getWb.Close
'
'                    Else
'                    MsgBox "File does not exist"
'                End If


            End If
        Next r
    Next cl

End Sub



Sub RevertCol()

    For Each sht In ActiveWorkbook.Worksheets
    Set rng = sht.UsedRange
    Set MyRange = rng
    
    Columns("A:A").Select
Selection.ClearContents
Columns("B:B").Select
Selection.ClearContents
                      
        For Each MyCol In MyRange.Columns
            For Each myCell In MyCol.Cells
            
            
            
Rows(1).Interior.Color = xlNone
Rows(1).Font.Color = 1

            
                If myCell.Interior.ColorIndex = 3 Then
               myCell.Interior.ColorIndex = xlNone
                
               End If
                  
            If myCell.Interior.Color = 8 Then
            myCell.Interior.ColorIndex = xlNone
                End If
               
            Next
        Next
    Next

End Sub




Sub myFileopener()

Dim myfile As String
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


myfile = Dir(myPath & "*.xls")
Do While myfile <> ""


If Not myfile Like "*_??.*" Then





Workbooks.Open (myPath & myfile)






          '  Workbooks.Open (File)
            ' Put code to process files here.
            
            
            MsgBox myfile
            
        End If

myfile = Dir


Loop

End Sub




