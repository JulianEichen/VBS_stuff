' VerschiebeDatei.vbs
' Eine Datei verschieben
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, DateiNameQuelle, DateiNameZiel
' Konstanten definieren
Const DateiNameQuelle="beispiel.txt"
Const DateiNameZiel="beispiel3.txt"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(DateiNameQuelle) Then
    ' Verschieben mit MoveFile
    FSO.MoveFile DateiNameQuelle, DateiNameZiel
    ' Ausgabe
    WScript.Echo DateiNameQuelle & " wurde nach " & _
    DateiNameZiel & " verschoben."
Else
    ' Fehlermeldung ausgeben
    WScript.Echo DateiNameQuelle & " ist nicht vorhanden"
End If