' VerzeichnisAttribute.vbs
' Attribute eines Verzeichnisses
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, Verzeichnis
' Konstanten definieren
Const VerzeichnisName="INetPub"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Gibt es das Verzeichnis überhaupt?
if FSO.FolderExists(VerzeichnisName) then
    ' Ja, also eine Verbindung herstellen
    Set Verzeichnis = FSO.GetFolder(VerzeichnisName)
    WScript.Echo "Typ des Objekts : " & Verzeichnis.Type
    WScript.Echo "Elternverzeichnis : " & Verzeichnis.ParentFolder
    WScript.Echo "ShortName : " & Verzeichnis.ShortName
    WScript.Echo "Erstellt am : " & Verzeichnis.DateCreated
    WScript.Echo "Geändert am : " & _
    Verzeichnis.DateLastModified 
    WScript.Echo "Letzter Zugriff : " & _
    Verzeichnis.DateLastAccessed 
    WScript.Echo "Attribute des Objekts : " & Verzeichnis.Attributes
    WScript.Echo "-----------------------"
    If Verzeichnis.Attributes AND 2 Then
        WScript.Echo "Versteckter Ordner"
    End if
    If Verzeichnis.Attributes AND 4 Then
        WScript.Echo "Systemordner"
    End if
    If Verzeichnis.Attributes AND 16 Then
        WScript.Echo "Ordner"
    End if
    If Verzeichnis.Attributes AND 32 Then
        WScript.Echo "Archive Bit gesetzt"
    End if
    If Verzeichnis.Attributes AND 2048 Then
        WScript.Echo "Komprimierter Ordner"
    End if
else
    WScript.Echo "Verzeichnis " & VerzeichnisName & " nicht gefunden!"
end if