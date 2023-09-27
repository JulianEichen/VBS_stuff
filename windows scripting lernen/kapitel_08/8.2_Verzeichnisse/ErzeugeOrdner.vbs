' ErzeugeOrdner.vbs
' Erzeugen eines Verzeichnisses
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, Verzeichnis
' Konstanten definieren
Const VerzeichnisName="Test"
' Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Prüfung, ob das Verzeichnis bereits existiert
if Not FSO.FolderExists(VerzeichnisName) then
    ' Verzeichnis anlegen
    FSO.CreateFolder(VerzeichnisName)
else
    ' Fehlermeldung ausgeben
    WScript.Echo "Verzeichnis " & VerzeichnisName & " existiert bereits"
End If