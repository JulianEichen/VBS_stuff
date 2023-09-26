' DateiUmbenennen.vbs
' Umbenennen einer Datei
' verwendet: SCRRun
' ==================================================
Option Explicit
' Deklaration der Variablen
Dim Dateisystem, Datei
' Konstanten definieren
Const DateiPfadAlt="c:\ausgabe.xml"
Const DateiNameNeu="Skriptausgabe.xml"
'FSO-Objekt erzeugen
Set DateiSystem = CreateObject("Scripting.FileSystemObject")
'File-Objekt gewinnen
Set Datei = Dateisystem.GetFile(DateiPfadAlt)
'Neuen Namen setzen
Datei.Name = DateiNameNeu
'Erfolgsmeldung ausgeben
MsgBox "Datei wurde umbenannt!"