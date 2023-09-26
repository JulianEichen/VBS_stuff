' LoescheDatei.vbs
' Löschen von Dateien
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, Datei, Verzeichnis
' FSO erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Referenz auf Verzeichnis holen
Set Verzeichnis = FSO.GetFolder("c:\inetpub")
' Alle Dateien bearbeiten
For Each Datei In Verzeichnis.Files
    ' Wenn Dateiendung .WMF dann
    If UCase(Right(Datei.Name, 4)) = ".WMF" Then
        ' Ausgabe
        WScript.Echo "Loesche " & Datei.Name
        ' Lösche Datei
        Datei.Delete
    End If
Next