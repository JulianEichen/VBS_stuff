' CSVDatei_Lesen.vbs
' Lesen einer CSV-Datei und Ausgabe der gelesenen Werte
' verwendet: SCRRun
' =========================================================================
Option Explicit
' Konstanten definieren
Const ForReading = 1
Const Dateiname="Benutzerliste.csv"
' Variablen deklarieren
Dim FSO, Datei, Benutzer
Dim TextZeile
'Objekt erzeugen
Set FSO=CreateObject("Scripting.FileSystemObject")
'Öffnen der Datei zum Lesen
Set Datei = FSO.OpenTextFile(Dateiname, ForReading, False)
'Datei bis zum Ende durchlaufen
While not Datei.AtEndOfStream
'Lesen einer Zeile
TextZeile=Datei.Readline()
'Zeile an Semikolon trennen und die Werte
'in einem Array speichern
Benutzer=Split(TextZeile,";")
'Ausgabe der Benutzerdaten
Wscript.echo Benutzer(0) & ";" & Benutzer(1) & ";" & Benutzer(2) & _
    ";" & Benutzer(3) & ";" & Benutzer(4)
Wend
'Schließen der Datei
Datei.Close