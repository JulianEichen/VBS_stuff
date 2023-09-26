' XMLDatei_Lesen.vbs
' Lesen einer XML-Datei und Ausgabe der Werte an der Konsole
' verwendet: XML
' ==========================================================
Option Explicit
' Deklaration der Variablen
Dim XMLDokument
Dim Benutzerknoten
Dim Zaehler
' Erzeugen des Verweises
Set XMLDokument = CreateObject("Msxml2.DOMDocument")
' Asynchrones Laden ausschalten
XMLDokument.async = False
' Datei laden
XMLDokument.load("Benutzer.xml")
' Knoten-Auflistung auswählen
Set Benutzerknoten = XMLDokument.selectNodes("*/*/USER")
' Alle Knoten durchlaufen
For Zaehler=0 to Benutzerknoten.length-1
    ' Daten ausgeben
    Wscript.echo Benutzerknoten.item(Zaehler).childNodes.item(0).Text & ";" _
    & Benutzerknoten.item(Zaehler).childNodes.item(1).Text & ";" _
    & Benutzerknoten.item(Zaehler).childNodes.item(2).Text & ";" & _
    Benutzerknoten.item(Zaehler).childNodes.item(3).Text & ";" & _
    Benutzerknoten.item(Zaehler).childNodes.item(4).Text
Next