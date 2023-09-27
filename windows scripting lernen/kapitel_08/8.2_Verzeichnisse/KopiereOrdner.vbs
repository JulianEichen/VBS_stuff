' KopiereOrdner.vbs
' Kopieren eines Verzeichnisses
' verwendet: SCRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO
' Konstanten definieren
Const VerzeichnisNameQuelle="Test"
Const VerzeichnisNameZiel="Test1"
' FSO-Objekt erstellen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Zielordner bereits vorhanden?
if not FSO.FolderExists(VerzeichnisNameZiel) then
    ' Quellverzeichnis vorhanden?
    if FSO.FolderExists(VerzeichnisNameQuelle) then
        ' Kopieren des Ordners
        FSO.CopyFolder VerzeichnisNameQuelle,VerzeichnisNameZiel
        WScript.echo "Ordner " & VerzeichnisNameQuelle & " wurde nach " & _
        VerzeichnisNameZiel & " kopiert"
    else
        WScript.echo "Quellordner " & VerzeichnisNameQuelle & _
        " existiert nicht"
    end if
else
     WScript.echo "Zielordner " & VerzeichnisNameZiel & " existiert bereits"
end if