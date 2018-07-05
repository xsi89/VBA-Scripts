Attribute VB_Name = "rws_XsiSec_140819"
Sub RWS()

'   RWS Makro
'   XsiSec RWS Makro
'   Detta Makro �r skrivet av XsiSec senast uppdaterat 140819
'   Det G�r f�ljande saker:
' * Kopierar Claimsdelen ur source filer f�r kunden RWS
' * tar det bort tomma rader i det inklistrade NewEuropat.dot fr�n sourcetexten
' * N�r det g�ller formatet p� text s� sker f�ljande:
' * Times new Roman, Teckenstorlek 12, radavst�nd 1.5, Tar bort all Fetstil
' * s�tter Br�dtext och marginaljusterad
' * �ppnar databas h�mtar jobbnummer plockar ut EPS nummer och klistrar in i sidhuvud
' * den g�r in p� webbsidan EPO h�mtar rubriken och klistrar in
' * Skapar Mappen translation to i jobbnumret
' * Kopierar Pdf:r till translation to
' * Sparar dokument med r�tt namn i translation to
'--------------------------------------------------------------------------------------------------------------------------------------------------

    Dim conn As ADODB.Connection
    Dim rst, rst2 As ADODB.recordset
    Dim sConnString As String
    Dim fileCount As Integer
    Dim StrJobnrValue As String
    
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))

    fileCount = 1
    sConnString = "Provider=SQLOLEDB;Data Source=SERVER05\SQLEXPRESS;" & _
                  "Initial Catalog=SQLJobbBackEnd1;" & _
                  "Integrated Security=SSPI;"

    Set conn = New ADODB.Connection
    conn.Open sConnString


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Spr�k.Spr�kkort_sv FROM (Spr�kparDK INNER JOIN JobbDK ON Spr�kparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Spr�k ON Spr�kparDK.Spr�knr = Spr�k.Spr�knr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

    myLang = rst!Spr�kkort_sv ' detta �r tabellen/cellen den h�mtar i DBN  till exempel en,den eller fra
    MsgBox myLang

    End If
    rst.Close
    conn.Close
    Set conn = Nothing
    
'-------------------------------------------------------ovanst�ende kod kopplar datorbasen --------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------



Select Case myLang

        Case "en"
            Call RWS_engelska
        Case "de"
            Call RWS_tyska
        Case "fra"
            Call RWS_franska
        Case Else
            MsgBox "Det �r n�got annat spr�k"
    End Select
'-------------------------------------- ovanst�ende �r loopen som kollar vilket rws skript som skall k�ras -------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------


End Sub



Sub RWS_engelska()
'-------------------------------------------------------------- H�r kopplas datorbasen ----------------------------------------------------------
    Dim conn As ADODB.Connection
    Dim rst, rst2 As ADODB.recordset
    Dim sConnString As String
    Dim fileCount As Integer
    Dim StrJobnrValue As String
    
    fileCount = 1
    sConnString = "Provider=SQLOLEDB;Data Source=SERVER05\SQLEXPRESS;" & _
                  "Initial Catalog=SQLJobbBackEnd1;" & _
                  "Integrated Security=SSPI;"
    Set conn = New ADODB.Connection
    conn.Open sConnString
'nedanst�ende kod filtrerar ut jobbnumret ur s�kv�gen -------------------------------------------------------------------------------------------
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))
'StrJobnrValue = 741477
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Spr�k.Spr�kkort_sv FROM (Spr�kparDK INNER JOIN JobbDK ON Spr�kparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Spr�k ON Spr�kparDK.Spr�knr2 = Spr�k.Spr�knr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then 'anledningen till det �r en "if not" �r att om den inte hittar jobbnumret s� f�r du inget felmeddelande
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

'MsgBox myEP

    End If ' h�r slutar if:n
    rst.Close
    conn.Close
    Set conn = Nothing
    
'--------------------------------------------------------ovanst�ende kod kopplar datorbasen ----------------------------------------------------
'----------------------- H�r h�mtas rubriken fr�n webbsidan skapar, �ven variabeln "InnovationTitle" ------------------------------------------

Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
     xmlHttp.Open "GET", "http://ops.epo.org/3.1/rest-services/published-data/publication/epodoc/" & myEP & "/biblio", False
    xmlHttp.setRequestHeader "Content-Type", "application/fulltext+xml"
    xmlHttp.Send
    If IsNull(xmlHttp.responsetext) Then
        errorText = "Kunde inte h�mta biblio-texten fr�n ops.epo.org. Meddelandet " & objMail.Subject & " har inte behandlats."
       
    End If
    biblioText = xmlHttp.responsetext
    myString = "<invention-title lang="
    myLength = Len(myString)
    myStart = InStr(biblioText, myString) + 5
    If (myStart = 5) Then
        errorText = "Texten '&lt;invention-title lang=' �terfanns inte i biblio-texten. Meddelandet har inte behandlats."
        MsgBox errorText
        
    End If
    myStop = InStr(myStart, biblioText, "<")
    innovationTitle = Trim(Mid(biblioText, myStart + myLength, myStop - myStart - myLength)) ' H�r deklareras variablen f�r rubriken InnovationTitle
        
    'MsgBox innovationTitle
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------
'---------------- Denna del kollar vad du �ppnat source filen,sedan s� kopierar den alla *.pdf filer d�r ----------------------------------------
'---------------- navigerar en niv� bak�t skapar mappen "translation to" om den finns s� l�gger den �nd� in *pdf:erna den kopierade i mappen ----

Dim StrOldPath As String, StrNewPath As String, strFile
StrOldPath = ActiveDocument.Path
StrNewPath = Left(StrOldPath, InStrRev(StrOldPath, "\"))
StrOldPath = StrOldPath & "\"
StrNewPath = StrNewPath & "translation to\"

If Dir(StrNewPath, vbDirectory) = "" Then
  MkDir StrNewPath
End If
strFile = Dir(StrOldPath & "*.PDF", vbNormal) '// v�ljer jag "*.pdf"
While strFile <> ""
  FileCopy StrOldPath & strFile, StrNewPath & strFile
  strFile = Dir()
Wend

'------------------------------------------------------------------------------------------------------------------------------------------------
'----------- Denna del s�ker efter ordet claims med tv� mjukreturer f�re och �ven till tv� mjukreturer efter + tv� radbrytningar ----------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^11^11Claims^013*^013^11^11"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    With Selection.Find
    Selection.Find.Execute
    Selection.Range.HighlightColorIndex = wdBrightGreen
    End With
    
'-----------------------------------------------------------------------------------------------------------------------------------------------
' h�r s�tter jag bland annat activedokument p� det dokument jag �ppnat, �ven s� �ppnar jag mallen "NewEuropat.dot" som ligger p� serven(G:/patent)
Dim thisdoc As Document
Dim Str
Set thisdoc = ActiveDocument
Dim targDoc As Word.Document
Set targDoc = Application.Documents.Open("G:\patent\NewEuropat.dot")
Documents(thisdoc).Activate

For Each Str In ActiveDocument.StoryRanges
    Str.Find.ClearFormatting
    Str.Find.Text = ""
    Str.Find.Highlight = True

While Str.Find.Execute
    Str.Copy
    Documents(targDoc).Activate
    Selection.MoveRight Unit:=wdCharacter, count:=50
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.Paste
    Documents(thisdoc).Activate
Wend
Next

'-----------------------------------------------------------------------------------------------------------------------------------------------
'------------------ H�r best�mmer formatet p� all text i NewEuropat.dot dit den kopierat texten fr�n sourcefilen -------------------------------
    With cleanform
        Documents(targDoc).Activate ' aktiverar NewEuropat.dot
        Selection.WholeStory '
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight ' tar bort all highlightext
        Selection.Font.Bold = wdToggle
        Selection.Font.Bold = wdToggle
        Selection.Font.Name = "Times New Roman" ' s�tter fonten
        Selection.Font.Size = 12 ' s�tter storlek
            
    End With
          With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5 ' s�tter radavst�nd
            .Alignment = wdAlignParagraphJustify
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            End With
'-------------------------------------------------------------------------------------------------------------------------------------------------
 '----------------------- h�r letar den efter det som st�r f�r rubriken sedan ers�tter den med en variabel ---------------------------------------
            
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Translated title to be inserted at the top of the page" ' Hitta denna text och ers�tt med variabeln nedanst�ende
        .Replacement.Text = (innovationTitle) ' h�r �r variablen f�r rubriken
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    
' -----------------------------------------------------------------------------------------------------------------------------------------------
'------------- H�r b�rjar scriptet jobba med sidhuvudet d� den st�ller mark�ren p� r�tt st�lle och erst�tter en del text(??????) ----------------
          
If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.MoveRight Unit:=wdCharacter, count:=7
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.TypeText Text:=(myEP)
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.EscapeKey

'-----------------------------------------------------------------------------------------------------------------------------------------------
'------------------- h�r sparar den filen och �ven best�mmer r�ttformatinst�llningar och anger r�tt s�kv�g. ------------------------------------
    ActiveDocument.SaveAs2 FileName:= _
        "H:\Jobb\RWS\" & (StrJobnrValue) & "\translation to\\" & (myEPSnum) & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14

End Sub

Sub RWS_tyska()
'-------------------------------------------------------------- H�r kopplas datorbasen ----------------------------------------------------------
    Dim conn As ADODB.Connection
    Dim rst, rst2 As ADODB.recordset
    Dim sConnString As String
    Dim fileCount As Integer
    Dim StrJobnrValue As String
    
    fileCount = 1
    sConnString = "Provider=SQLOLEDB;Data Source=SERVER05\SQLEXPRESS;" & _
                  "Initial Catalog=SQLJobbBackEnd1;" & _
                  "Integrated Security=SSPI;"
    Set conn = New ADODB.Connection
    conn.Open sConnString
'nedanst�ende kod filtrerar ut jobbnumret ur s�kv�gen -------------------------------------------------------------------------------------------
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))
'StrJobnrValue = 741477
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Spr�k.Spr�kkort_sv FROM (Spr�kparDK INNER JOIN JobbDK ON Spr�kparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Spr�k ON Spr�kparDK.Spr�knr2 = Spr�k.Spr�knr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then 'anledningen till det �r en "if not" �r att om den inte hittar jobbnumret s� f�r du inget felmeddelande
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

'MsgBox myEP

    End If ' h�r slutar if:n
    rst.Close
    conn.Close
    Set conn = Nothing
    
'--------------------------------------------------------ovanst�ende kod kopplar datorbasen ----------------------------------------------------
'----------------------- H�r h�mtas rubriken fr�n webbsidan skapar, �ven variabeln "InnovationTitle" ------------------------------------------

Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
     xmlHttp.Open "GET", "http://ops.epo.org/3.1/rest-services/published-data/publication/epodoc/" & myEP & "/biblio", False
    xmlHttp.setRequestHeader "Content-Type", "application/fulltext+xml"
    xmlHttp.Send
    If IsNull(xmlHttp.responsetext) Then
        errorText = "Kunde inte h�mta biblio-texten fr�n ops.epo.org. Meddelandet " & objMail.Subject & " har inte behandlats."
       
    End If
    biblioText = xmlHttp.responsetext
    myString = "<invention-title lang="
    myLength = Len(myString)
    myStart = InStr(biblioText, myString) + 5
    If (myStart = 5) Then
        errorText = "Texten '&lt;invention-title lang=' �terfanns inte i biblio-texten. Meddelandet har inte behandlats."
        MsgBox errorText
        
    End If
    myStop = InStr(myStart, biblioText, "<")
    innovationTitle = Trim(Mid(biblioText, myStart + myLength, myStop - myStart - myLength)) ' H�r deklareras variablen f�r rubriken InnovationTitle
        
    'MsgBox innovationTitle
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------
'---------------- Denna del kollar vad du �ppnat source filen,sedan s� kopierar den alla *.pdf filer d�r ----------------------------------------
'---------------- navigerar en niv� bak�t skapar mappen "translation to" om den finns s� l�gger den �nd� in *pdf:erna den kopierade i mappen ----

Dim StrOldPath As String, StrNewPath As String, strFile
StrOldPath = ActiveDocument.Path
StrNewPath = Left(StrOldPath, InStrRev(StrOldPath, "\"))
StrOldPath = StrOldPath & "\"
StrNewPath = StrNewPath & "translation to\"

If Dir(StrNewPath, vbDirectory) = "" Then
  MkDir StrNewPath
End If
strFile = Dir(StrOldPath & "*.PDF", vbNormal) '// v�ljer jag "*.pdf"
While strFile <> ""
  FileCopy StrOldPath & strFile, StrNewPath & strFile
  strFile = Dir()
Wend

'------------------------------------------------------------------------------------------------------------------------------------------------
'----------- Denna del s�ker efter ordet claims med tv� mjukreturer f�re och �ven till tv� mjukreturer efter + tv� radbrytningar ----------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^11^11Anspr�che^013*^013^11^11"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    With Selection.Find
    Selection.Find.Execute
    Selection.Range.HighlightColorIndex = wdBrightGreen
    End With
    
'-----------------------------------------------------------------------------------------------------------------------------------------------
' h�r s�tter jag bland annat activedokument p� det dokument jag �ppnat, �ven s� �ppnar jag mallen "NewEuropat.dot" som ligger p� serven(G:/patent)
Dim thisdoc As Document
Dim Str
Set thisdoc = ActiveDocument
Dim targDoc As Word.Document
Set targDoc = Application.Documents.Open("G:\patent\NewEuropat.dot")
Documents(thisdoc).Activate

For Each Str In ActiveDocument.StoryRanges
    Str.Find.ClearFormatting
    Str.Find.Text = ""
    Str.Find.Highlight = True

While Str.Find.Execute
    Str.Copy
    Documents(targDoc).Activate
    Selection.MoveRight Unit:=wdCharacter, count:=50
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.Paste
    Documents(thisdoc).Activate
Wend
Next

'-----------------------------------------------------------------------------------------------------------------------------------------------
'------------------ H�r best�mmer formatet p� all text i NewEuropat.dot dit den kopierat texten fr�n sourcefilen -------------------------------
    With cleanform
        Documents(targDoc).Activate ' aktiverar NewEuropat.dot
        Selection.WholeStory '
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight ' tar bort all highlightext
        Selection.Font.Bold = wdToggle
        Selection.Font.Bold = wdToggle
        Selection.Font.Name = "Times New Roman" ' s�tter fonten
        Selection.Font.Size = 12 ' s�tter storlek
            
    End With
          With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5 ' s�tter radavst�nd
            .Alignment = wdAlignParagraphJustify
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            End With
'-------------------------------------------------------------------------------------------------------------------------------------------------
 '----------------------- h�r letar den efter det som st�r f�r rubriken sedan ers�tter den med en variabel ---------------------------------------
            
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Translated title to be inserted at the top of the page" ' Hitta denna text och ers�tt med variabeln nedanst�ende
        .Replacement.Text = (innovationTitle) ' h�r �r variablen f�r rubriken
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    
' -----------------------------------------------------------------------------------------------------------------------------------------------
'------------- H�r b�rjar scriptet jobba med sidhuvudet d� den st�ller mark�ren p� r�tt st�lle och erst�tter en del text(??????) ----------------
          
If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.MoveRight Unit:=wdCharacter, count:=7
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.TypeText Text:=(myEP)
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.EscapeKey

'-----------------------------------------------------------------------------------------------------------------------------------------------
'------------------- h�r sparar den filen och �ven best�mmer r�ttformatinst�llningar och anger r�tt s�kv�g. ------------------------------------
    ActiveDocument.SaveAs2 FileName:= _
        "H:\Jobb\RWS\" & (StrJobnrValue) & "\translation to\\" & (myEPSnum) & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14

End Sub
        

Sub RWS_franska()
'-------------------------------------------------------------- H�r kopplas datorbasen ----------------------------------------------------------
    Dim conn As ADODB.Connection
    Dim rst, rst2 As ADODB.recordset
    Dim sConnString As String
    Dim fileCount As Integer
    Dim StrJobnrValue As String
    
    fileCount = 1
    sConnString = "Provider=SQLOLEDB;Data Source=SERVER05\SQLEXPRESS;" & _
                  "Initial Catalog=SQLJobbBackEnd1;" & _
                  "Integrated Security=SSPI;"
    Set conn = New ADODB.Connection
    conn.Open sConnString
'nedanst�ende kod filtrerar ut jobbnumret ur s�kv�gen -------------------------------------------------------------------------------------------
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))
'StrJobnrValue = 741477
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Spr�k.Spr�kkort_sv FROM (Spr�kparDK INNER JOIN JobbDK ON Spr�kparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Spr�k ON Spr�kparDK.Spr�knr2 = Spr�k.Spr�knr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then 'anledningen till det �r en "if not" �r att om den inte hittar jobbnumret s� f�r du inget felmeddelande
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

'MsgBox myEP

    End If ' h�r slutar if:n
    rst.Close
    conn.Close
    Set conn = Nothing
    
'--------------------------------------------------------ovanst�ende kod kopplar datorbasen ----------------------------------------------------
'----------------------- H�r h�mtas rubriken fr�n webbsidan skapar, �ven variabeln "InnovationTitle" ------------------------------------------

Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
     xmlHttp.Open "GET", "http://ops.epo.org/3.1/rest-services/published-data/publication/epodoc/" & myEP & "/biblio", False
    xmlHttp.setRequestHeader "Content-Type", "application/fulltext+xml"
    xmlHttp.Send
    If IsNull(xmlHttp.responsetext) Then
        errorText = "Kunde inte h�mta biblio-texten fr�n ops.epo.org. Meddelandet " & objMail.Subject & " har inte behandlats."
       
    End If
    biblioText = xmlHttp.responsetext
    myString = "<invention-title lang="
    myLength = Len(myString)
    myStart = InStr(biblioText, myString) + 5
    If (myStart = 5) Then
        errorText = "Texten '&lt;invention-title lang=' �terfanns inte i biblio-texten. Meddelandet har inte behandlats."
        MsgBox errorText
        
    End If
    myStop = InStr(myStart, biblioText, "<")
    innovationTitle = Trim(Mid(biblioText, myStart + myLength, myStop - myStart - myLength)) ' H�r deklareras variablen f�r rubriken InnovationTitle
        
    'MsgBox innovationTitle
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------
'---------------- Denna del kollar vad du �ppnat source filen,sedan s� kopierar den alla *.pdf filer d�r ----------------------------------------
'---------------- navigerar en niv� bak�t skapar mappen "translation to" om den finns s� l�gger den �nd� in *pdf:erna den kopierade i mappen ----

Dim StrOldPath As String, StrNewPath As String, strFile
StrOldPath = ActiveDocument.Path
StrNewPath = Left(StrOldPath, InStrRev(StrOldPath, "\"))
StrOldPath = StrOldPath & "\"
StrNewPath = StrNewPath & "translation to\"

If Dir(StrNewPath, vbDirectory) = "" Then
  MkDir StrNewPath
End If
strFile = Dir(StrOldPath & "*.PDF", vbNormal) '// v�ljer jag "*.pdf"
While strFile <> ""
  FileCopy StrOldPath & strFile, StrNewPath & strFile
  strFile = Dir()
Wend

'------------------------------------------------------------------------------------------------------------------------------------------------
'----------- Denna del s�ker efter ordet claims med tv� mjukreturer f�re och �ven till tv� mjukreturer efter + tv� radbrytningar ----------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^11^11Revendications^013*^013^11^11"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    With Selection.Find
    Selection.Find.Execute
    Selection.Range.HighlightColorIndex = wdBrightGreen
    End With
    
'-----------------------------------------------------------------------------------------------------------------------------------------------
' h�r s�tter jag bland annat activedokument p� det dokument jag �ppnat, �ven s� �ppnar jag mallen "NewEuropat.dot" som ligger p� serven(G:/patent)
Dim thisdoc As Document
Dim Str
Set thisdoc = ActiveDocument
Dim targDoc As Word.Document
Set targDoc = Application.Documents.Open("G:\patent\NewEuropat.dot")
Documents(thisdoc).Activate

For Each Str In ActiveDocument.StoryRanges
    Str.Find.ClearFormatting
    Str.Find.Text = ""
    Str.Find.Highlight = True

While Str.Find.Execute
    Str.Copy
    Documents(targDoc).Activate
    Selection.MoveRight Unit:=wdCharacter, count:=50
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.Paste
    Documents(thisdoc).Activate
Wend
Next

'-----------------------------------------------------------------------------------------------------------------------------------------------
'------------------ H�r best�mmer formatet p� all text i NewEuropat.dot dit den kopierat texten fr�n sourcefilen -------------------------------
    With cleanform
        Documents(targDoc).Activate ' aktiverar NewEuropat.dot
        Selection.WholeStory '
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight ' tar bort all highlightext
        Selection.Font.Bold = wdToggle
        Selection.Font.Bold = wdToggle
        Selection.Font.Name = "Times New Roman" ' s�tter fonten
        Selection.Font.Size = 12 ' s�tter storlek
            
    End With
          With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5 ' s�tter radavst�nd
            .Alignment = wdAlignParagraphJustify
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            End With
'-------------------------------------------------------------------------------------------------------------------------------------------------
 '----------------------- h�r letar den efter det som st�r f�r rubriken sedan ers�tter den med en variabel ---------------------------------------
            
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Translated title to be inserted at the top of the page" ' Hitta denna text och ers�tt med variabeln nedanst�ende
        .Replacement.Text = (innovationTitle) ' h�r �r variablen f�r rubriken
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    
' -----------------------------------------------------------------------------------------------------------------------------------------------
'------------- H�r b�rjar scriptet jobba med sidhuvudet d� den st�ller mark�ren p� r�tt st�lle och erst�tter en del text(??????) ----------------
          
If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.MoveRight Unit:=wdCharacter, count:=7
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.TypeText Text:=(myEP)
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.EscapeKey

'-----------------------------------------------------------------------------------------------------------------------------------------------
'------------------- h�r sparar den filen och �ven best�mmer r�ttformatinst�llningar och anger r�tt s�kv�g. ------------------------------------
    ActiveDocument.SaveAs2 FileName:= _
        "H:\Jobb\RWS\" & (StrJobnrValue) & "\translation to\\" & (myEPSnum) & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14

        End Sub
