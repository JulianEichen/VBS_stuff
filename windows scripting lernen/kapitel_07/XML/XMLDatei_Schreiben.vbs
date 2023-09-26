' XMLDatei_Schreiben.vbs
' Lesen einer CSV-Datei und Ausgabe in ein XML-Dokument
' verwendet: XML, SCRRun
' ===============================
Option Explicit
' Konstanten definieren
Const ForReading = 1
' Variablen deklarieren
Dim FSO, Datei, Benutzer
Dim TextZeile
Dim XMLDokument, XMLWurzel, XMLBenutzer, XMLBenutzerliste
Dim XMLBenutzername, XMLVorname, XMLNachname, XMLGeburtsdatum, XMLAbteilungsnummer
'Objekt erzeugen
Set FSO=CreateObject("Scripting.FileSystemObject")
'Öffnen der Datei zum Lesen
Set Datei = FSO.OpenTextFile("Benutzerliste.csv", ForReading, False)
' Erzeugen des XML-Dokuments
Set XMLDokument = CreateObject("Msxml.DOMDocument")
' Erzeugen des Wurzelelements
Set XMLWurzel = XMLDokument.createElement("BENUTZER")
' Hinzufügen des Wurzelelements zum Dokument
XMLDokument.appendChild XMLWurzel
' Erzeugen der Benutzerliste - Auflistung
Set XMLBenutzerliste = XMLDokument.createElement("USERS")
XMLWurzel.appendChild XMLBenutzerliste
' Überlesen der Feldnamen
Datei.SkipLine()
'Datei bis zum Ende durchlaufen
while not Datei.AtEndOfStream
    ' Erzeugen der Benutzer-Collection
    Set XMLBenutzer=XMLDokument.createElement("USER")
    XMLBenutzerliste.appendChild XMLBenutzer
    'Lesen einer Zeile
    TextZeile=Datei.Readline()
    'Zeile an Semikolon trennen und die Werte
    'in einem Array speichern
    Benutzer=Split(TextZeile,";")
    'Ausgabe der Benutzerdaten in ein XML-Dokument
    Set XMLBenutzername = XMLDokument.createElement("BENUTZERNAME")
    XMLBenutzername.Text = Benutzer(0)
    XMLBenutzer.appendChild XMLBenutzername
    Set XMLVorname = XMLDokument.createElement("VORNAME")
    XMLVorname.Text = Benutzer(1)
    XMLBenutzer.appendChild XMLVorname
    Set XMLNachname = XMLDokument.createElement("NACHNAME")
    XMLNachname.Text = Benutzer(2)
    XMLBenutzer.appendChild XMLNachname
    Set XMLGeburtsdatum = XMLDokument.createElement("GEBURTSDATUM")
    XMLGeburtsdatum.Text = Benutzer(3)
    XMLBenutzer.appendChild XMLGeburtsdatum
    Set XMLAbteilungsnummer = XMLDokument.createElement("ABTEILUNGSNUMMER")
    XMLAbteilungsnummer.Text = Benutzer(4)
    XMLBenutzer.appendChild XMLAbteilungsnummer
Wend
' XML-Datei speichern
XMLDokument.Save "Benutzer.xml"
'Schließen der Datei
Datei.Close