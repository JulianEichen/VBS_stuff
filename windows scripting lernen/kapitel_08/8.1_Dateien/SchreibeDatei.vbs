' SchreibeDatei.vbsClose
' Eine Textdatei schreiben
' verwendet: SCRRun
' ===============================
Option Explicit
'Konstantendefinitionen
Const ForWriting = 2
Const DateiName="beispiel.txt"
' Deklaration der Variablen
Dim FSO, DateiInhalt, Zaehler
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
set DateiInhalt = FSO.OpenTextFile(DateiName, ForWriting)
'Alle Buchstaben des Alphabets in die Datei schreiben
For Zaehler = 1 To 26
    'Kleinbuchstaben beginnen an Position 97
    DateiInhalt.Write Chr(96 + Zaehler)
    'Großbuchstaben beginnen an Position 65
    DateiInhalt.Write Chr(64 + Zaehler)
Next
'Datei schließen
DateiInhalt.Close