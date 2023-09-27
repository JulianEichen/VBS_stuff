' LeseOrdner.vbs
' Lesen eines Verzeichnisses
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, Verzeichnis, UnterVerzeichnis
Dim Datei
' Konstanten definieren
Const VerzeichnisName="INetPub"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
'Referenz auf ein Verzeichnis holen
Set Verzeichnis = FSO.GetFolder(VerzeichnisName)
' Ausgabe
WScript.Echo "-- Dateien:"
' Alle Dateien
For Each Datei In Verzeichnis.Files
    WScript.Echo Datei.Name
Next
WScript.Echo "-- Ordner:"
' Alle Unterverzeichnisse
For Each UnterVerzeichnis In Verzeichnis.SubFolders
    WScript.Echo UnterVerzeichnis.Name
Next