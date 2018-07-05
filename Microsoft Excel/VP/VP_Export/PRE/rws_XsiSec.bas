Attribute VB_Name = "rws_XsiSec_140819"
Sub RWS()

'   RWS Makro
'   XsiSec RWS Makro
'   Detta Makro är skrivet av XsiSec senast uppdaterat 140819
'   Det Gör följande saker:
' * Kopierar Claimsdelen ur source filer för kunden RWS
' * tar det bort tomma rader i det inklistrade NewEuropat.dot från sourcetexten
' * När det gäller formatet på text så sker följande:
' * Times new Roman, Teckenstorlek 12, radavstånd 1.5, Tar bort all Fetstil
' * sätter Brödtext och marginaljusterad
' * öppnar databas hämtar jobbnummer plockar ut EPS nummer och klistrar in i sidhuvud
' * den går in på webbsidan EPO hämtar rubriken och klistrar in
' * Skapar Mappen translation to i jobbnumret
' * Kopierar Pdf:r till translation to
' * Sparar dokument med rätt namn i translation to
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


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Språk.Språkkort_sv FROM (SpråkparDK INNER JOIN JobbDK ON SpråkparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Språk ON SpråkparDK.Språknr = Språk.Språknr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

    myLang = rst!Språkkort_sv ' detta är tabellen/cellen den hämtar i DBN  till exempel en,den eller fra
    MsgBox myLang

    End If
    rst.Close
    conn.Close
    Set conn = Nothing
    
'-------------------------------------------------------ovanstående kod kopplar datorbasen --------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------



Select Case myLang

        Case "en"
            Call RWS_engelska
        Case "de"
            Call RWS_tyska
        Case "fra"
            Call RWS_franska
        Case Else
            MsgBox "Det är något annat språk"
    End Select
'-------------------------------------- ovanstående är loopen som kollar vilket rws skript som skall köras -------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------


End Sub



Sub RWS_engelska()
'-------------------------------------------------------------- Här kopplas datorbasen ----------------------------------------------------------
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
'nedanstående kod filtrerar ut jobbnumret ur sökvägen -------------------------------------------------------------------------------------------
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))
'StrJobnrValue = 741477
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Språk.Språkkort_sv FROM (SpråkparDK INNER JOIN JobbDK ON SpråkparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Språk ON SpråkparDK.Språknr2 = Språk.Språknr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then 'anledningen till det är en "if not" är att om den inte hittar jobbnumret så får du inget felmeddelande
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

'MsgBox myEP

    End If ' här slutar if:n
    rst.Close
    conn.Close
    Set conn = Nothing
    
'--------------------------------------------------------ovanstående kod kopplar datorbasen ----------------------------------------------------
'----------------------- Här hämtas rubriken från webbsidan skapar, även variabeln "InnovationTitle" ------------------------------------------

Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
     xmlHttp.Open "GET", "http://ops.epo.org/3.1/rest-services/published-data/publication/epodoc/" & myEP & "/biblio", False
    xmlHttp.setRequestHeader "Content-Type", "application/fulltext+xml"
    xmlHttp.Send
    If IsNull(xmlHttp.responsetext) Then
        errorText = "Kunde inte hämta biblio-texten från ops.epo.org. Meddelandet " & objMail.Subject & " har inte behandlats."
       
    End If
    biblioText = xmlHttp.responsetext
    myString = "<invention-title lang="
    myLength = Len(myString)
    myStart = InStr(biblioText, myString) + 5
    If (myStart = 5) Then
        errorText = "Texten '&lt;invention-title lang=' återfanns inte i biblio-texten. Meddelandet har inte behandlats."
        MsgBox errorText
        
    End If
    myStop = InStr(myStart, biblioText, "<")
    innovationTitle = Trim(Mid(biblioText, myStart + myLength, myStop - myStart - myLength)) ' Här deklareras variablen för rubriken InnovationTitle
        
    'MsgBox innovationTitle
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------
'---------------- Denna del kollar vad du öppnat source filen,sedan så kopierar den alla *.pdf filer där ----------------------------------------
'---------------- navigerar en nivå bakåt skapar mappen "translation to" om den finns så lägger den ändå in *pdf:erna den kopierade i mappen ----

Dim StrOldPath As String, StrNewPath As String, strFile
StrOldPath = ActiveDocument.Path
StrNewPath = Left(StrOldPath, InStrRev(StrOldPath, "\"))
StrOldPath = StrOldPath & "\"
StrNewPath = StrNewPath & "translation to\"

If Dir(StrNewPath, vbDirectory) = "" Then
  MkDir StrNewPath
End If
strFile = Dir(StrOldPath & "*.PDF", vbNormal) '// väljer jag "*.pdf"
While strFile <> ""
  FileCopy StrOldPath & strFile, StrNewPath & strFile
  strFile = Dir()
Wend

'------------------------------------------------------------------------------------------------------------------------------------------------
'----------- Denna del söker efter ordet claims med två mjukreturer före och även till två mjukreturer efter + två radbrytningar ----------------
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
' här sätter jag bland annat activedokument på det dokument jag öppnat, även så öppnar jag mallen "NewEuropat.dot" som ligger på serven(G:/patent)
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
'------------------ Här bestämmer formatet på all text i NewEuropat.dot dit den kopierat texten från sourcefilen -------------------------------
    With cleanform
        Documents(targDoc).Activate ' aktiverar NewEuropat.dot
        Selection.WholeStory '
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight ' tar bort all highlightext
        Selection.Font.Bold = wdToggle
        Selection.Font.Bold = wdToggle
        Selection.Font.Name = "Times New Roman" ' sätter fonten
        Selection.Font.Size = 12 ' sätter storlek
            
    End With
          With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5 ' sätter radavstånd
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
 '----------------------- här letar den efter det som står för rubriken sedan ersätter den med en variabel ---------------------------------------
            
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Translated title to be inserted at the top of the page" ' Hitta denna text och ersätt med variabeln nedanstående
        .Replacement.Text = (innovationTitle) ' här är variablen för rubriken
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
'------------- Här börjar scriptet jobba med sidhuvudet då den ställer markören på rätt ställe och erstätter en del text(??????) ----------------
          
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
'------------------- här sparar den filen och även bestämmer rättformatinställningar och anger rätt sökväg. ------------------------------------
    ActiveDocument.SaveAs2 FileName:= _
        "H:\Jobb\RWS\" & (StrJobnrValue) & "\translation to\\" & (myEPSnum) & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14

End Sub

Sub RWS_tyska()
'-------------------------------------------------------------- Här kopplas datorbasen ----------------------------------------------------------
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
'nedanstående kod filtrerar ut jobbnumret ur sökvägen -------------------------------------------------------------------------------------------
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))
'StrJobnrValue = 741477
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Språk.Språkkort_sv FROM (SpråkparDK INNER JOIN JobbDK ON SpråkparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Språk ON SpråkparDK.Språknr2 = Språk.Språknr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then 'anledningen till det är en "if not" är att om den inte hittar jobbnumret så får du inget felmeddelande
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

'MsgBox myEP

    End If ' här slutar if:n
    rst.Close
    conn.Close
    Set conn = Nothing
    
'--------------------------------------------------------ovanstående kod kopplar datorbasen ----------------------------------------------------
'----------------------- Här hämtas rubriken från webbsidan skapar, även variabeln "InnovationTitle" ------------------------------------------

Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
     xmlHttp.Open "GET", "http://ops.epo.org/3.1/rest-services/published-data/publication/epodoc/" & myEP & "/biblio", False
    xmlHttp.setRequestHeader "Content-Type", "application/fulltext+xml"
    xmlHttp.Send
    If IsNull(xmlHttp.responsetext) Then
        errorText = "Kunde inte hämta biblio-texten från ops.epo.org. Meddelandet " & objMail.Subject & " har inte behandlats."
       
    End If
    biblioText = xmlHttp.responsetext
    myString = "<invention-title lang="
    myLength = Len(myString)
    myStart = InStr(biblioText, myString) + 5
    If (myStart = 5) Then
        errorText = "Texten '&lt;invention-title lang=' återfanns inte i biblio-texten. Meddelandet har inte behandlats."
        MsgBox errorText
        
    End If
    myStop = InStr(myStart, biblioText, "<")
    innovationTitle = Trim(Mid(biblioText, myStart + myLength, myStop - myStart - myLength)) ' Här deklareras variablen för rubriken InnovationTitle
        
    'MsgBox innovationTitle
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------
'---------------- Denna del kollar vad du öppnat source filen,sedan så kopierar den alla *.pdf filer där ----------------------------------------
'---------------- navigerar en nivå bakåt skapar mappen "translation to" om den finns så lägger den ändå in *pdf:erna den kopierade i mappen ----

Dim StrOldPath As String, StrNewPath As String, strFile
StrOldPath = ActiveDocument.Path
StrNewPath = Left(StrOldPath, InStrRev(StrOldPath, "\"))
StrOldPath = StrOldPath & "\"
StrNewPath = StrNewPath & "translation to\"

If Dir(StrNewPath, vbDirectory) = "" Then
  MkDir StrNewPath
End If
strFile = Dir(StrOldPath & "*.PDF", vbNormal) '// väljer jag "*.pdf"
While strFile <> ""
  FileCopy StrOldPath & strFile, StrNewPath & strFile
  strFile = Dir()
Wend

'------------------------------------------------------------------------------------------------------------------------------------------------
'----------- Denna del söker efter ordet claims med två mjukreturer före och även till två mjukreturer efter + två radbrytningar ----------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^11^11Ansprüche^013*^013^11^11"
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
' här sätter jag bland annat activedokument på det dokument jag öppnat, även så öppnar jag mallen "NewEuropat.dot" som ligger på serven(G:/patent)
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
'------------------ Här bestämmer formatet på all text i NewEuropat.dot dit den kopierat texten från sourcefilen -------------------------------
    With cleanform
        Documents(targDoc).Activate ' aktiverar NewEuropat.dot
        Selection.WholeStory '
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight ' tar bort all highlightext
        Selection.Font.Bold = wdToggle
        Selection.Font.Bold = wdToggle
        Selection.Font.Name = "Times New Roman" ' sätter fonten
        Selection.Font.Size = 12 ' sätter storlek
            
    End With
          With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5 ' sätter radavstånd
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
 '----------------------- här letar den efter det som står för rubriken sedan ersätter den med en variabel ---------------------------------------
            
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Translated title to be inserted at the top of the page" ' Hitta denna text och ersätt med variabeln nedanstående
        .Replacement.Text = (innovationTitle) ' här är variablen för rubriken
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
'------------- Här börjar scriptet jobba med sidhuvudet då den ställer markören på rätt ställe och erstätter en del text(??????) ----------------
          
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
'------------------- här sparar den filen och även bestämmer rättformatinställningar och anger rätt sökväg. ------------------------------------
    ActiveDocument.SaveAs2 FileName:= _
        "H:\Jobb\RWS\" & (StrJobnrValue) & "\translation to\\" & (myEPSnum) & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14

End Sub
        

Sub RWS_franska()
'-------------------------------------------------------------- Här kopplas datorbasen ----------------------------------------------------------
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
'nedanstående kod filtrerar ut jobbnumret ur sökvägen -------------------------------------------------------------------------------------------
Filepath = ActiveDocument.Path
myPath = ActiveDocument.Path
Filepath = Split(myPath, "\")
StrJobnrValue = (Filepath(3))
'StrJobnrValue = 741477
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------


    Set rst = conn.Execute("SELECT JobbDK.JOBBNR, JobbDK.JOBBESKR, Språk.Språkkort_sv FROM (SpråkparDK INNER JOIN JobbDK ON SpråkparDK.Jobbnr = JobbDK.JOBBNR) INNER JOIN Språk ON SpråkparDK.Språknr2 = Språk.Språknr WHERE (((JobbDK.JOBBNR)=" & (StrJobnrValue) & "));")
    If Not rst.EOF Then 'anledningen till det är en "if not" är att om den inte hittar jobbnumret så får du inget felmeddelande
    myJobbnr = rst!JOBBNR
    myEPSnum = rst!JOBBESKR
    myEPSnum = Replace(myEPSnum, " ", "")
    myIndex = InStr(1, myEPSnum, "EP")
    myEP = Mid(myEPSnum, myIndex, 9)

'MsgBox myEP

    End If ' här slutar if:n
    rst.Close
    conn.Close
    Set conn = Nothing
    
'--------------------------------------------------------ovanstående kod kopplar datorbasen ----------------------------------------------------
'----------------------- Här hämtas rubriken från webbsidan skapar, även variabeln "InnovationTitle" ------------------------------------------

Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
     xmlHttp.Open "GET", "http://ops.epo.org/3.1/rest-services/published-data/publication/epodoc/" & myEP & "/biblio", False
    xmlHttp.setRequestHeader "Content-Type", "application/fulltext+xml"
    xmlHttp.Send
    If IsNull(xmlHttp.responsetext) Then
        errorText = "Kunde inte hämta biblio-texten från ops.epo.org. Meddelandet " & objMail.Subject & " har inte behandlats."
       
    End If
    biblioText = xmlHttp.responsetext
    myString = "<invention-title lang="
    myLength = Len(myString)
    myStart = InStr(biblioText, myString) + 5
    If (myStart = 5) Then
        errorText = "Texten '&lt;invention-title lang=' återfanns inte i biblio-texten. Meddelandet har inte behandlats."
        MsgBox errorText
        
    End If
    myStop = InStr(myStart, biblioText, "<")
    innovationTitle = Trim(Mid(biblioText, myStart + myLength, myStop - myStart - myLength)) ' Här deklareras variablen för rubriken InnovationTitle
        
    'MsgBox innovationTitle
'--- SLUT ---------------------------------------------------------------------------------------------------------------------------------------
'---------------- Denna del kollar vad du öppnat source filen,sedan så kopierar den alla *.pdf filer där ----------------------------------------
'---------------- navigerar en nivå bakåt skapar mappen "translation to" om den finns så lägger den ändå in *pdf:erna den kopierade i mappen ----

Dim StrOldPath As String, StrNewPath As String, strFile
StrOldPath = ActiveDocument.Path
StrNewPath = Left(StrOldPath, InStrRev(StrOldPath, "\"))
StrOldPath = StrOldPath & "\"
StrNewPath = StrNewPath & "translation to\"

If Dir(StrNewPath, vbDirectory) = "" Then
  MkDir StrNewPath
End If
strFile = Dir(StrOldPath & "*.PDF", vbNormal) '// väljer jag "*.pdf"
While strFile <> ""
  FileCopy StrOldPath & strFile, StrNewPath & strFile
  strFile = Dir()
Wend

'------------------------------------------------------------------------------------------------------------------------------------------------
'----------- Denna del söker efter ordet claims med två mjukreturer före och även till två mjukreturer efter + två radbrytningar ----------------
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
' här sätter jag bland annat activedokument på det dokument jag öppnat, även så öppnar jag mallen "NewEuropat.dot" som ligger på serven(G:/patent)
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
'------------------ Här bestämmer formatet på all text i NewEuropat.dot dit den kopierat texten från sourcefilen -------------------------------
    With cleanform
        Documents(targDoc).Activate ' aktiverar NewEuropat.dot
        Selection.WholeStory '
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight ' tar bort all highlightext
        Selection.Font.Bold = wdToggle
        Selection.Font.Bold = wdToggle
        Selection.Font.Name = "Times New Roman" ' sätter fonten
        Selection.Font.Size = 12 ' sätter storlek
            
    End With
          With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5 ' sätter radavstånd
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
 '----------------------- här letar den efter det som står för rubriken sedan ersätter den med en variabel ---------------------------------------
            
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "Translated title to be inserted at the top of the page" ' Hitta denna text och ersätt med variabeln nedanstående
        .Replacement.Text = (innovationTitle) ' här är variablen för rubriken
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
'------------- Här börjar scriptet jobba med sidhuvudet då den ställer markören på rätt ställe och erstätter en del text(??????) ----------------
          
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
'------------------- här sparar den filen och även bestämmer rättformatinställningar och anger rätt sökväg. ------------------------------------
    ActiveDocument.SaveAs2 FileName:= _
        "H:\Jobb\RWS\" & (StrJobnrValue) & "\translation to\\" & (myEPSnum) & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14

        End Sub
