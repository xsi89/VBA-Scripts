VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "MT Feedback Report                                                       ® Copyright Teknotrans AB - Written by Daniel Elmnäs"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14550
   OleObjectBlob   =   "Form3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub myrun()


End Sub


Private Sub CommandButton1_Click()

Call DelLasRow

End Sub

 Sub UserForm_Initialize()



  For i = 0 To 300
        For j = 1 To 5
            Me.Time.AddItem i & "," & j
        Next j
        Me.Time.AddItem i + 1
    Next i

With Me.PerC
    .AddItem "10 %"
    .AddItem "20 %"
    .AddItem "30 %"
    .AddItem "40 %"
    .AddItem "50 %"
    .AddItem "60 %"
    .AddItem "70 %"
    .AddItem "80 %"
    .AddItem "90 %"
    .AddItem "100 %"
End With

With Me.SourLang
.AddItem " ar-SA "
.AddItem " en-AU "
.AddItem " bg-BG "
.AddItem " bs-Latn-BA "
.AddItem " pt-BR "
.AddItem " fr-CA "
.AddItem " zh-HK "
.AddItem " zh-CN "
.AddItem " hr-HR "
.AddItem " cs-CZ "
.AddItem " da-DK "
.AddItem " de-DE "
.AddItem " en-GB "
.AddItem " es-ES "
.AddItem " et-EE "
.AddItem " fi-FI "
.AddItem " nl-BE "
.AddItem " fr-FR "
.AddItem " el-GR "
.AddItem " he-IL "
.AddItem " hi-IN "
.AddItem " hu-HU "
.AddItem " id-ID "
.AddItem " fa-IR "
.AddItem " is-IS "
.AddItem " it-IT "
.AddItem " ja-JP "
.AddItem " ko-KR "
.AddItem " lt-LT "
.AddItem " lv-LV "
.AddItem " mk-MK "
.AddItem " es-MX "
.AddItem " ms-MY "
.AddItem " nl-NL "
End With

With Me.TarLang
.AddItem " ar-SA "
.AddItem " en-AU "
.AddItem " bg-BG "
.AddItem " bs-Latn-BA "
.AddItem " pt-BR "
.AddItem " fr-CA "
.AddItem " zh-HK "
.AddItem " zh-CN "
.AddItem " hr-HR "
.AddItem " cs-CZ "
.AddItem " da-DK "
.AddItem " de-DE "
.AddItem " en-GB "
.AddItem " es-ES "
.AddItem " et-EE "
.AddItem " fi-FI "
.AddItem " nl-BE "
.AddItem " fr-FR "
.AddItem " el-GR "
.AddItem " he-IL "
.AddItem " hi-IN "
.AddItem " hu-HU "
.AddItem " id-ID "
.AddItem " fa-IR "
.AddItem " is-IS "
.AddItem " it-IT "
.AddItem " ja-JP "
.AddItem " ko-KR "
.AddItem " lt-LT "
.AddItem " lv-LV "
.AddItem " mk-MK "
.AddItem " es-MX "
.AddItem " ms-MY "
.AddItem " nl-NL "
End With

With TypeOfMis
.AddItem "Consistency"
.AddItem "Grammar"
.AddItem "Mistranslation"
.AddItem "Sentence structure"
.AddItem "Terminology"
.AddItem "Other (please specify in comments)"
End With




With Info4

Info4.Caption = "The TT (target text) is syntactically correct, it uses proper terminology, it conveys information accurately and uses an appropriate style. Your understanding is not improved by the reading of the ST (source text)." & vbCrLf & "Effect: no corrections are required."
End With

With Info3

Info3.Caption = "Your understanding is not improved by the reading of the ST even though the post edited segment contains minor errors affecting any of these: grammatical (article, preposition), syntax (word order), punctuation, word formation (verb endings, number agreement), inappropriate style. An end user who does not have access to the source text could anyway understand the post edited segments." & vbCrLf & "Effect: Only a few corrections required in terms of actual changes or time spent."

End With

With Info2

Info2.Caption = "Your understanding is improved by the reading of the ST, due to significant errors in the post edited segments. You would have to re-read the ST a few times to correct these errors in the post edited segment. An end user who does not have access to the source text could only get a general understanding of the post edited segments' meaning. " & vbCrLf & "Effect: Severe post-editing is required or maybe just minor post-editing after spending too much time trying to understand the intended meaning and where the errors are."


End With

With Info1

Info1.Caption = "Your understanding only derives from reading the ST, as you could not understand the post edited segment since it contained serious errors. You could only produce a proofread translation by dismissing most of the post edited segment and/or re-translating from scratch. An end user who does not have access to the source text would not be able to understand the post edited segment at all." & vbCrLf & "Effect: It would be better to manually re-translate from scratch (proofreading of the post-editing is not worthwhile). "

End With

With Info0

Info0.Caption = "PE: post-editing" & vbCrLf & "MT raw output: the translation provided by the MT BEFORE post-editing"

End With




End Sub



Private Sub cmdSend_Click()
'Range("A3").Value = "G" + Proj.Text 'Project
'Range("B3").Value = SourLang.Text ' Source language
'Range("C3").Value = TarLang.Text ' Target Language
'Range("F3").Value = TypeOfMis.Text ' Type of Mistakes Listbox
'Range("I3").Value = Cmnts.Text ' Comments
'Range("H3").Value = PerC.Text ' Percent is ok
'Range("G3").Value = ExaOfMis.Text ' example of mistakes
Range("B" & Rows.Count).End(xlUp).Offset(1).Value = Proj.Text
Range("C" & Rows.Count).End(xlUp).Offset(1).Value = SourLang.Text
Range("D" & Rows.Count).End(xlUp).Offset(1).Value = TarLang.Text

Range("E" & Rows.Count).End(xlUp).Offset(1).Value = Fuzzy.Text
Range("F" & Rows.Count).End(xlUp).Offset(1).Value = NewOfWords.Text
Range("G" & Rows.Count).End(xlUp).Offset(1).Value = repetitions.Text

Range("H" & Rows.Count).End(xlUp).Offset(1).Value = Time.Text
Range("I" & Rows.Count).End(xlUp).Offset(1).Value = TypeOfMis.Text
Range("J" & Rows.Count).End(xlUp).Offset(1).Value = ExaOfMis.Text
Range("K" & Rows.Count).End(xlUp).Offset(1).Value = PerC.Text
Range("L" & Rows.Count).End(xlUp).Offset(1).Value = Cmnts.Text

Call ConvNum



End Sub






