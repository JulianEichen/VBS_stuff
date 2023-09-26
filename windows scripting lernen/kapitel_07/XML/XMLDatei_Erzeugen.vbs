' XMLDatei_Erzeugen.vbs
' Erzeugen einer XML-Datei
' verwendet: XML
' ===============================
Option Explicit
' Variablen deklarieren
Dim XMLDokument
Dim XMLWurzel
' XML-Dokument erzeugen
Set XMLDokument = CreateObject("Msxml.DOMDocument")
' Wurzelelement erzeugen
Set XMLWurzel = XMLDokument.CreateElement("BENUTZER")
' Wurzelelement an das Dokument anhängen
XMLDokument.AppendChild XMLWurzel
' Datei speichern
XMLDokument.Save "User.xml"