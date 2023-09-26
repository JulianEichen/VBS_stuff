' LeseDateiZeilenweise.vbs
' Eine Textdatei zeilenweise lesen
' verwendet: SCRRun
' ==========================================================
Option Explicit
' Deklaration der Variablen
Dim FSO, DateiInhalt, Zeile, Inhalt
' Konstanten definieren
Const Dateiname="beispiel.txt"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Gibt es die Datei überhaupt?
If FSO.FileExists(DateiName) Then
    ' Ja, also eine Verbindung herstellen
    Set DateiInhalt = FSO.OpenTextFile(Dateiname)
    Do Until DateiInhalt.atEndOfStream
        Zeile = DateiInhalt.ReadLine
        Inhalt=Inhalt + Zeile + vbcrlf
    Loop
    DateiInhalt.Close
    WScript.Echo Inhalt
    Else
        WScript.Echo "Datei " & Dateiname & " nicht gefunden!"
End If