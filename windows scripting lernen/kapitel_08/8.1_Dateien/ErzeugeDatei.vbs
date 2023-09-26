' ErzeugeDatei.vbs
' Erzeugen einer Textdatei
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO
Dim Datei
' Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
'Datei Beispiel.txt erzeugen
Set Datei = Fso.CreateTextFile("Beispiel.txt", True)
'Eine Zeile in die Datei schreiben
Datei.WriteLine("Dies ist meine erste automatisch erzeugte Datei")
'Datei schließen
Datei.Close