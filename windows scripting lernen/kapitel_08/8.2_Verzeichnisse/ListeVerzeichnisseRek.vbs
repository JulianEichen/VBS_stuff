' ListeVerzeichnisseRek.vbs
' Auflisten von Verzeichnissen
' verwendet: SCRRun
' ===============================
Option Explicit
' Aufruf der Routine
' Konstanten definieren
Const VerzeichnisBezeichner="."
ListeVerzeichnisseRek VerzeichnisBezeichner
Sub ListeVerzeichnisseRek(Verzeichnisname)
' Deklaration der Variablen
Dim FSO, Verzeichnis, UnterVerzeichnis
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Wenn das Verzeichnis existiert
if FSO.FolderExists(Verzeichnisname) then
    ' Ordner holen
    Set Verzeichnis = FSO.GetFolder(Verzeichnisname)
    ' Alle Unterverzeichnisse auflisten
    for each UnterVerzeichnis in Verzeichnis.subfolders
        WScript.Echo UnterVerzeichnis.Name
        ' Erneuter Aufruf mit dem Unterverzeichnis
        ListeVerzeichnisseRek UnterVerzeichnis
    next
end if
End Sub