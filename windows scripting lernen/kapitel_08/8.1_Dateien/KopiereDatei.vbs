' KopiereDatei.vbs
' Eine Datei kopieren
' verwendet: SCRRun
' =========================================
Option Explicit
' Deklaration der Variablen
Dim FSO
' Konstanten definieren
Const DateiNameQuelle="beispiel.txt"
Const DateiNameZiel="beispiel2.txt"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(DateiNameQuelle) Then
    'Kopieren mit CopyFile
    FSO.CopyFile DateiNameQuelle, DateiNameZiel, True
    WScript.Echo DateiNameQuelle & " wurde nach " & _
    DateiNameZiel & " kopiert."
Else
    WScript.Echo DateiNameQuelle & " ist nicht vorhanden."
End If