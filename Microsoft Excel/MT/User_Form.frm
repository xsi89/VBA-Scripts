VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "MT Post Editing Feedback"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   OleObjectBlob   =   "User_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()



    For i = 0.1 To 28 Step 0.1
        Me.Time.AddItem Format(i, "0.0")
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

End Sub



Private Sub cmdSend_Click()
Range("A3").Value = "G" + Proj.Text
Range("B3").Value = SourLang.Text
Range("C3").Value = TarLang.Text
Range("F3").Value = TypeOfMis.Text
Range("I3").Value = Cmnts.Text
Range("H3").Value = PerC.Text

End Sub


