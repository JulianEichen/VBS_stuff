' DateienAuflisten.vbs
' Auflisten von Dateien
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Konstanten
Const Verzeichnis = "C:\temp"
' Deklaration der Variablen
Dim FSO, Verzeichnis, Datei
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
'Referenz auf ein Verzeichnis holen
Set Verzeichnis = FSO.GetFolder(Verzeichnis)
'Über alle Dateien im Verzeichnis iterieren
For Each Datei In Verzeichnis.Files
    WScript.Echo Datei.Name
Next