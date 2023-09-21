' Fehlerbehandlung.vbs
' Abfangen von Fehlermeldungen
' verwendet: keine weiteren Komponenten
' =============================== 
' Ordentliche Fehlerbehandlung
Sub Fehlermeldung
    If Err.Number<>0 Then
    WScript.Echo("Es ist ein Fehler aufgetreten")
    WScript.Echo(vbTab & "Fehlernummer: " & Err.Number)
    WScript.Echo(vbTab & "Beschreibung: " & Err.Description & vbCrLf)
    Err.Description
 End If
End Sub
Dim wert
' Deaktivieren von Fehlerabbrüchen
On Error Resume Next
' Division durch Null
wert = 2 / 0
Fehlermeldung
' Aktivierung des Fehlerabbruchs
On Error GoTo 0
' Division durch Null
wert = 2 / 0
Fehlermeldung