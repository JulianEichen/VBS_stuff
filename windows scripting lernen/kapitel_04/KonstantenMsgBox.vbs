' KonstantenMsgBox.vbs
' Konstanten für die Verwendung in Nachrichtenfenstern
' verwendet: keine weiteren Komponenten
' =========================================================
Option Explicit

Sub WelcherButtonWurdeGedrueckt(tmpButtonKonstante)
    Dim strButtonName
    Select Case tmpButtonKonstante
        Case vbOk
            strButtonName = "Ok"
        Case vbCancel
            strButtonName = "Abbrechen"
        Case vbAbort
            strButtonName = "Abbrechen"
        Case vbRetry
            strButtonName = "Wiederholen"
        Case vbIgnore
            strButtonName = "Ignorieren"
        Case vbYes
            strButtonName = "Ja"
        Case vbNo
            strButtonName = "Nein"
        Case vbElse
            strButtonName = "Unbekannter Button"
    End Select

    MsgBox "Es wurde der " & strButtonName & _
        "-Button gedrückt", vbInformation, "Information"
End Sub

Dim tmpWert

' vbOkCancel + vbCritical 
tmpWert = MsgBox("Achtung! Hier ist ein Stoppschild, also bitte anhalten!", _
vbOkCancel + vbCritical, "Stopp-Symbol")
WelcherButtonWurdeGedrueckt(tmpWert)

' vbAbortRetryIgnore + vbQuestion
tmpWert = MsgBox("Dies ist eine Frage", vbAbortRetryIgnore + vbQuestion, _
"Fragezeichen-Symbol")
WelcherButtonWurdeGedrueckt(tmpWert)
' vbYesNoCancel + vbExclamation
tmpWert = MsgBox("Achtung! Dies ist eine Warnung", vbYesNoCancel + _
vbExclamation + vbDefaultButton3, "Warnung-Symbol")
WelcherButtonWurdeGedrueckt(tmpWert)
' vbRetryCancel + vbInformation 
tmpWert = MsgBox("Dies ist nur eine Information", vbRetryCancel + _
vbInformation + vbDefaultButton2, "Information-Symbol")
WelcherButtonWurdeGedrueckt(tmpWert)