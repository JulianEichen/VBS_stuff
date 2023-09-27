' LoescheVerzeichnis.vbs
' Leeren des Papierkorbs für alle Benutzer
' verwendet: Shell.Application
' ===============================
Set DateiSystem = CreateObject("Scripting.FileSystemObject")
' Zugriff auf Papierkorb
Set objFolder = DateiSystem.GetFolder("c:\$recycle.bin\")
' Unterordner im Papierkorb holen
Set Ordnerliste = objFolder.SubFolders
' Schleife über alle Elemente mit Fallunterscheidung zwischen Ordner und Datei
For Each Element In Ordnerliste
    WScript.Echo "Lösche Ordner: " + Element.Path
    On Error Resume Next
    DateiSystem.DeleteFolder Element.Path, True
    If Err.number > 0 Then
        WScript.Echo "Fehler: " & Err.Description
        Err.Clear
    End If
Next