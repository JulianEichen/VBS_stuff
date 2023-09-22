' CSVDatei_Erzeugen.vbs 
' Anlegen einer CSV-Datei und Speichern der Feldnamen
' verwendet: SCRRun
' ===========================================================================
Option Explicit
' Deklaration der Variablen
Dim FSO,Datei
Const Dateiname="Benutzerliste.csv"
'Erzeugen einer Objektreferenz
Set FSO = CreateObject("Scripting.FileSystemObject")
'Erzeugen der Datei
Set Datei = FSO.CreateTextFile(Dateiname)
' Schreiben der Spaltennamen
Datei.WriteLine("Benutzername;Vorname;Nachname;Geburtstag;Abteilungsnummer")
Datei.Close