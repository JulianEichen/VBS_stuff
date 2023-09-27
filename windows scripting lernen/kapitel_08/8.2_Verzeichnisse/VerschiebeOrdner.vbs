' VerschiebeOrdner.vbs
' Verschieben eines Verzeichnisses
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO
' Konstanten definieren
Const VerzeichnisNameQuelle="Test"
Const VerzeichnisNameZiel="Test1"
' FSO-Objekt erstellen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Wenn die Quelle existiert, dann
if FSO.FolderExists(VerzeichnisNameQuelle) then
    ' Verschieben des Ordners
    FSO.MoveFolder VerzeichnisNameQuelle,VerzeichnisNameZiel
    WScript.Echo "Ordner " & VerzeichnisNameQuelle & _
    " wurde nach " & VerzeichnisNameZiel & " verschoben"
else
    WScript.Echo "Quelle " & VerzeichnisNameQuelle & " existiert nicht"
end if