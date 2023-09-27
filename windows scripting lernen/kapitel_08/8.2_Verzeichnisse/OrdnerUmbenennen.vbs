' OrdnerUmbenennen.vbs
' Umbenennen eines Dateisystemordners
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim Dateisystem, Ordner
' Konstanten definieren
Const OrdnerPfadAlt="c:\test"
Const OrdnerNameNeu="Skripte"
'FSO-Objekt erzeugen
Set DateiSystem = CreateObject("Scripting.FileSystemObject")
'File-Objekt gewinnen
Set Ordner = Dateisystem.GetFolder(OrdnerPfadAlt)
'Neuen Namen setzen
Ordner.Name = OrdnerNameNeu
'Erfolgsmeldung ausgeben
MsgBox "Ordner wurde umbenannt!"