Attribute VB_Name = "VolvoPenta_Import"
Sub LoopFiles()

Dim myPath As String
Dim StrCurrentfile As String
Dim StrFName As String



StrDocPAth = "C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org - Copy\"
StrCurrentfile = Dir(StrDocPAth & "*NoTrans.xls")
Do While StrCurrentfile <> ""


myNoTransfile = StrDocPAth & StrCurrentfile

myLangFile = Replace(StrCurrentfile, "_NoTrans", "")






MsgBox myLangFile


MsgBox myNoTransfile



'strDocPath & StrCurrentfile

'Workbooks.Open strDocPath & StrCurrentfile



StrCurrentfile = Dir






Loop

End Sub


Sub PastebetweenColumnF()



Set SourceFile = Workbooks.Open("C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org\1\UCHP_Translation2_jeeves_sv_NoTrans.xls")
Columns(1).Select
Selection.Copy


Set TargetFile = Workbooks.Open("C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org\1\UCHP_Translation2_jeeves_sv.xls")
TargetFile.Sheets.Add(After:=TargetFile.Sheets(TargetFile.Sheets.Count)).Name = "WordNotTrans"

  ActiveSheet.Paste



End Sub



'The following code will combine all data into one excel workbook.
Sub CombineFiles_Step1()
'Declare Variables
Dim WorkbookDestination As Workbook
Dim WorkbookSource As Workbook
Dim WorksheetSource As Worksheet
Dim FolderLocation As String
Dim strFileName As String
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'This line will need to be modified depending on location of source folder
    FolderLocation = "C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org"
    
    'Set the current directory to the the folder path.
    ChDrive FolderLocation
    ChDir FolderLocation
    
    'Dialog box to determine which files to use. Use ctrl+a to select all files in folder.
'    SelectedFiles = Application.GetOpenFilename( _
'        filefilter:="Excel Files (*.xls*), *.xls*", MultiSelect:=True)



    
    'Create a new workbook
    Set WorkbookDestination = Workbooks.Add(xlWBATWorksheet)
    strFileName = Dir(FolderLocation & "\*.xls", vbNormal)
    
    'Iterate for each file in folder
    If Len(strFileName) = 0 Then Exit Sub
    
    
    Do Until strFileName = ""
        
            Set WorkbookSource = Workbooks.Open(Filename:=FolderLocation & "\" & strFileName)
            Set WorksheetSource = WorkbookSource.Worksheets(1)
            WorksheetSource.Copy After:=WorkbookDestination.Worksheets(WorkbookDestination.Worksheets.Count)
            WorkbookSource.Close False
        strFileName = Dir()
        
    Loop
    WorkbookDestination.Worksheets(1).Delete
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub



Sub Contentpaste()


Set SourceFile = Workbooks.Open("C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org\1\UCHP_Translation2_jeeves_sv_NoTrans.xls")
Columns(1).Select
Selection.Copy

Set TargetFile = Workbooks.Open("C:\Users\daniel.elmnas.TT\Desktop\ko\FIle\Org\1\UCHP_Translation2_jeeves_sv.xls")
TargetFile.Sheets.Add(After:=TargetFile.Sheets(TargetFile.Sheets.Count)).Name = "WordNotTrans"
ActiveSheet.Paste


ActiveWorkbook.Worksheets("Translated").Activate
Rows("1:1").Select
Selection.EntireRow.Hidden = False

Application.DisplayAlerts = False
SourceFile.Close SaveChanges:=False
End Sub





