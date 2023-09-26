' VerschiebeDatei2.vbs
' Eine Datei sicher verschieben
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
    ' Kopiere die Datei
    FSO.CopyFile DateiNameQuelle, DateiNameZiel, True
    ' Nun löschen, auch schreibgeschützt
    FSO.DeleteFile DateiNameQuelle, true
    ' Ausgabe
    WScript.Echo DateiNameQuelle & " wurde nach " & _
    DateiNameZiel & " verschoben."
Else
    ' Fehlermeldung ausgeben
    WScript.Echo DateiNameQuelle & " ist nicht vorhanden"
End If