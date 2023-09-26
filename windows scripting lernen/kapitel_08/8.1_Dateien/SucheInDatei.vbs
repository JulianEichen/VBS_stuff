' SucheInDatei.vbs
' Suchen in Dateien
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim Verzeichnis, Unterverzeichnis
Dim SuchText, FSO
' Suchtext aus der Kommandozeile lesen
SuchText = WScript.Arguments(0)
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Zu durchsuchendes Verzeichnis aus der Kommandozeile lesen
Set Verzeichnis = FSO.GetFolder(WScript.Arguments(1))
' Aufruf der Suchfunktion
WScript.Echo "Der Text " & SuchText & " wurde gefunden in:"
Suche Verzeichnis,SuchText

Function Suche(Verzeichnis,SucheText)
    Dim Dateien,TextStream,Dateiinhalt
    For Each Dateien in Verzeichnis.Files
        Set TextStream = FSO.OpenTextFile(Dateien.Path,1)
        Dateiinhalt = TextStream.ReadAll
        If InStr(1, Dateiinhalt, SucheText, 1) then
            WScript.Echo Dateien.Path
        End If
        TextStream.Close
    Next
    ' Unterverzeichnis durchsuchen
    For Each Unterverzeichnis in Verzeichnis.SubFolders
        Suche Unterverzeichnis,SucheText
    Next
End Function