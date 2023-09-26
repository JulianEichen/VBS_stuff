' LeseDateiKomplett.vbs
' Eine Textdatei komplett lesen
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, DateiInhalt, Zeile, Inhalt
' Konstanten definieren
Const DateiName="beispiel.txt"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Gibt es die Datei überhaupt?
if FSO.FileExists(DateiName) then
    ' Ja, also eine Verbindung herstellen
    set DateiInhalt = FSO.OpenTextFile(DateiName)
    'Gesamte Datei auf einmal lesen
    Inhalt=DateiInhalt.ReadAll()
    'Datei schließen
    DateiInhalt.Close
    'Inhalt ausgeben
    WScript.Echo Inhalt
else
    WScript.Echo "Datei " & DateiName& " nicht gefunden!"
end if