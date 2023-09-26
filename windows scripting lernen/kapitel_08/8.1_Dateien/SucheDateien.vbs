' SucheDateien.vbs
' Rekursives Suchen von Dateien
' verwendet: SCRRun
' Aufruf in Kommandozeile mit Parametern wie folgt:
' cscript.exe SucheDateien.vbs Startverzeichnis Suchwort
' ===============================
Option Explicit
Dim Start, Suchwort
If WScript.Arguments.Count = 2 Then
    ' Werte von der Kommandozeile lesen
    Start = WScript.Arguments(0)
    Suchwort=WScript.Arguments(1)
    ' Aufruf der Hilfsroutine
    ListeOrdner Start, Suchwort
else
    'Syntax ausgeben
    WScript.Echo "Syntax: SucheDateien.vbs Startverzeichnis Suchwort"
End If
' Hilfsroutine: Rekursion über Ordnerinhalte
Sub ListeOrdner(Ordner, Suchmaske)
' Deklaration der Variablen
Dim FSO, Verzeichnis, Datei, Unterverzeichnis
' Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Referenz auf Verzeichnis erzeugen
Set Verzeichnis = FSO.GetFolder(Ordner)
' Alle Dateien im Verzeichnis durchlaufen
For Each Datei In Verzeichnis.Files
    ' Wenn Dateiname mit Suchwort übereinstimmt
    If InStr(UCase(Datei.Name),UCase(Suchmaske))>0 Then
        ' Ausgabe des Pfades und des Dateinamens
        WScript.Echo "Gefunden: " & Datei.path
    End If
Next
' Durchlaufe alle Unterverzeichnisse
For Each Unterverzeichnis In Verzeichnis.SubFolders
    ' Wenn Verzeichnisname mit Suchwort übereinstimmt
    If InStr(UCase(Unterverzeichnis.Name),UCase(Suchmaske))>0 then
        ' Ausgabe des Pfades und des Verzeichnisnamens
        WScript.Echo Unterverzeichnis.Name
    End If
    'Rekursiver Aufruf der Routine
    ListeOrdner Unterverzeichnis.Path, Suchmaske
Next
End Sub