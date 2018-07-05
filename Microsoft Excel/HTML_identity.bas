Attribute VB_Name = "HTML_identity"
Sub Remove_HTML_Identities()

' detta script är skrivet av XsiSec 2016-03-18
' Sök och ersätter identiteter

Dim wb As Workbook
Dim myPath As String
Dim myfile As String
Dim myExtension As String
Dim FldrPicker As FileDialog


Dim LastRow, LastColumn As Long
Dim i, j As Long

Dim h1, r1 As String
Dim h2, r2 As String
Dim h3, r3 As String
Dim h4, r4 As String
Dim h5, r5 As String
Dim h6, r6 As String
Dim h7, r7 As String
Dim h8, r8 As String
Dim h9, r9 As String
Dim h10, r10 As String
Dim h11, r11 As String
Dim h12, r12 As String
Dim h13, r13 As String
Dim h14, r14 As String
Dim h15, r15 As String

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
  myExtension = "*.xls"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myfile)







h1 = "&auml;"
h2 = "&ouml;"
h3 = "&uuml;"
h4 = "&Auml;"
h5 = "&Ouml;"
h6 = "&Uuml;"
h7 = "&szlig;"
h8 = "&lsquo;"
h9 = "&rsquo;"
h10 = "&ldquo;"
h11 = "&rdquo;"
h12 = "&bdquo;"
h13 = "&Uuml;"
h14 = "&amp;"
h15 = "&#39;"

r1 = "ä"
r2 = "ö"
r3 = "ü"
r4 = "Ä"
r5 = "Ö"
r6 = "Ü"
r7 = "ß"
r8 = "‘"
r9 = "’"
r10 = "“"
r11 = "”"
r12 = "„"
r13 = "Ü"
r14 = "&"
r15 = "'"



LR = Range("A" & Rows.Count).End(xlUp).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column

For i = 1 To LR
For j = 1 To LC
 
      If Cells(i, j).EntireColumn.Hidden = False Then
        Cells(i, j).Replace what:=h1, Replacement:=r1, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h2, Replacement:=r2, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h3, Replacement:=r3, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h4, Replacement:=r4, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h5, Replacement:=r5, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h6, Replacement:=r6, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h7, Replacement:=r7, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h8, Replacement:=r8, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h9, Replacement:=r9, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h10, Replacement:=r10, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h11, Replacement:=r11, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h12, Replacement:=r12, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h13, Replacement:=r13, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h14, Replacement:=r14, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h15, Replacement:=r15, lookat:=xlPart, MatchCase:=False
    End If

Next j
NextRow:
Next i





      wb.Close SaveChanges:=True

    'Get next file name
      myfile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub



Sub replaceHTMLIdentites()
Dim LastRow, LastColumn As Long
Dim i, j As Long

Dim h1, r1 As String
Dim h2, r2 As String
Dim h3, r3 As String
Dim h4, r4 As String
Dim h5, r5 As String
Dim h6, r6 As String
Dim h7, r7 As String
Dim h8, r8 As String
Dim h9, r9 As String
Dim h10, r10 As String
Dim h11, r11 As String
Dim h12, r12 As String


h1 = "&auml;"
h2 = "&ouml;"
h3 = "&uuml;"
h4 = "&Auml;"
h5 = "&Ouml;"
h6 = "&Uuml;"
h7 = "&szlig;"
h8 = "&lsquo;"
h9 = "&rsquo;"
h10 = "&ldquo;"
h11 = "&rdquo;"
h12 = "&bdquo;"
h13 = "&Uuml;"
h14 = "&amp;"




r1 = "ä"
r2 = "ö"
r3 = "ü"
r4 = "Ä"
r5 = "Ö"
r6 = "Ü"
r7 = "ß"
r8 = "‘"
r9 = "’"
r10 = "“"
r11 = "”"
r12 = "„"
r13 = "Ü"
r14 = "&"





LR = Range("A" & Rows.Count).End(xlUp).Row
LC = Cells(1, Columns.Count).End(xlToLeft).Column

For i = 1 To LR
For j = 1 To LC
 
    If Cells(i, j).EntireColumn.Hidden = False Then
        Cells(i, j).Replace what:=h1, Replacement:=r1, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h2, Replacement:=r2, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h3, Replacement:=r3, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h4, Replacement:=r4, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h5, Replacement:=r5, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h6, Replacement:=r6, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h7, Replacement:=r7, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h8, Replacement:=r8, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h9, Replacement:=r9, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h10, Replacement:=r10, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h11, Replacement:=r11, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h12, Replacement:=r12, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h13, Replacement:=r13, lookat:=xlPart, MatchCase:=False
        Cells(i, j).Replace what:=h14, Replacement:=r14, lookat:=xlPart, MatchCase:=False
    End If

Next j
NextRow:
Next i
End Sub


Sub Find_Replace()

    Dim fndList     As Variant
    Dim rplcList    As Variant
    Dim x           As Long
    Dim LR As Long
    Dim LC As Long
    Dim r As Range
    
    

    
    
    fndList = Array("&auml;", "&ouml;", "&uuml;", "&Auml;", "&Ouml;", "&Uuml;", "&szlig;", "&lsquo;", "&rsquo;", "&ldquo;", "&rdquo;", "&bdquo;", "&Uuml;", "&amp;")
    rplcList = Array("ä", "ö", "ü", "Ä", "Ö", "Ü", "ß", "‘", "’", "“", "”", "„", "Ü", "&")
    Application.ScreenUpdating = False
    LR = Range("A" & Rows.Count).End(xlUp).Row
    LC = Cells(1, Columns.Count).End(xlToLeft).Column
        
    Set r = Range(Cells(1, 1), Cells(LR, LC)).SpecialCells(xlCellTypeVisible)
    
      For x = LBound(fndList) To UBound(fndList)
          r.Replace what:=fndList(x), Replacement:=rplcList(x), lookat:=xlPart
      Next x
    
    Application.ScreenUpdating = True
   
End Sub



