' LeseDateiZeichenweise.vbs
' Eine Textdatei zeichenweise lesen
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, DateiInhalt, Zeile, Inhalt, Zeichen
' Konstanten definieren
Const DateiName="beispiel.txt"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Gibt es die Datei überhaupt?
If FSO.FileExists(DateiName) then
    ' Ja, also eine Verbindung herstellen
    Set DateiInhalt = FSO.OpenTextFile(DateiName)
    'Solange das Ende der Datei nicht erreicht ist
    Do Until DateiInhalt.atEndOfStream
        'Ein Zeichen lesen
        Zeichen = DateiInhalt.Read(1)
        Inhalt=Inhalt + Zeichen + vbcrlf
    Loop
    'Datei schließen
    DateiInhalt.Close
    'Inhalt ausgeben
    WScript.Echo Inhalt
Else
    WScript.Echo "Datei " & DateiName & " nicht gefunden!"
End If