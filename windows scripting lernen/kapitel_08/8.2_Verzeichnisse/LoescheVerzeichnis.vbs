' LoescheVerzeichnis.vbs
' Löschen eines Verzeichnisses
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO
Const VerzeichnisName="w:\alteDokumente"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Wenn es das Verzeichnis gibt, dann ...
If FSO.FolderExists(VerzeichnisName) Then
    ' löschen
    FSO.DeleteFolder Verzeichnisname, true
    ' Ausgabe
    WScript.Echo "Das Verzeichnis " & VerzeichnisName & " wurde gelöscht."
Else
    ' sonst Fehlermeldung ausgeben
    WScript.Echo "Das Verzeichnis " & VerzeichnisName & _
    " existiert nicht."
End If