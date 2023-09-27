' VerzeichnisSuche.vbs
' Suchen eines Verzeichnisses (rekursiv)
' verwendet: SCRRun
' ===============================
Option Explicit
' Aufruf der Routine
SucheOrdner "C:\Winnt","System"
' === Unterroutine
Sub SucheOrdner(StartVerzeichnis,Suchtext)
    ' Deklaration der Variablen
    Dim FSO, Verzeichnis, Unterverzeichnis
    ' FSO-Objekt erstellen
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' Referenz auf Verzeichnis holen
    Set Verzeichnis = FSO.GetFolder(StartVerzeichnis)
    ' Durchlaufe Unterverzeichnisse
    For Each Unterverzeichnis In Verzeichnis.SubFolders
        ' Entspricht Verzeichnisname dem gesuchten Element?
        If InStr(UCase(Unterverzeichnis.Name),UCase(Suchtext))>0 then
            ' Ausgabe des Verzeichnisnamens
            WScript.Echo Unterverzeichnis.Name
        End If
            ' Rekursiver Aufruf für nächste Verzeichnisebene
            SucheOrdner Unterverzeichnis.Path,Suchtext
    Next
End Sub