' ------------------------------------------
' Skriptname: VerzeichnisstrukturDokumentieren.vbs
' ------------------------------------------
' Dieses Skript dokumentiert die Struktur
' eines Dateisystemverzeichnisses in XML-Form
' ------------------------------------------
' verwendet SCRRun, MSXML
' ------------------------------------------
Option Explicit
' Deklaration der Variablen
Dim FSO
Dim XMLDocument
Dim StartKnoten
Const MaxEbene = 3
Const Basisverzeichnis = "T:\"
Const Ausgabedatei = "T:\verzeichnisstruktur.xml"
' Notwendige COM-Objekte erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
Set XMLDocument = CreateObject("Msxml.DOMDocument")
' Prüfen, ob Basisverzeichnis vorhanden
If Not FSO.FolderExists(BasisVerzeichnis) Then
    WScript.Echo "Verzeichnis " & Basisverzeichnis & " existiert nicht!"
    WScript.Quit
End If
' Processing Instruction erzeugen
Dim pi
Set pi = XMLDocument.createProcessingInstruction("xml", " version=""1.0""")
XMLDocument.InsertBefore pi, XMLDocument.childNodes.Item(0)
' -- Erzeuge Root-Element
Dim Wurzel
Set Wurzel = xml_add(XMLDocument, XMLDocument, "VerzeichnisStruktur", "")
' Rekursion
VerzeichnisseDokumentieren Basisverzeichnis, Wurzel, 1
' Speichern
XMLDocument.save Ausgabedatei
WScript.Echo "Ausgabedatei wurde erfolgreich gespeichert!"
' === Rekursive Hilfsroutine zum Anlegen der XML-Knoten für jeden Ordner
Sub VerzeichnisseDokumentieren(Pfad, XmlKnoten,Ebene)
    Dim Ordner
    Dim Unterordner
    Dim ele
    ' Ordner holen
    Set Ordner = FSO.GetFolder(Pfad)
    WScript.Echo "Dokumentiere Ordner: " & Ordner.Path
    ' Element für Ordner erzeugen
    Set ele = xml_add(xmldocument, XmlKnoten, "Verzeichnis", "")
    ele.setAttribute "Name", Ordner.Name
    ' Maximale Dokumentationstiefe erreicht?
    if Ebene = MaxEbene then Exit Sub
        ' Schleife über die Unterordner
    For Each Unterordner In Ordner.SubFolders
        VerzeichnisseDokumentieren Unterordner.Path,ele, Ebene+1
    Next
End Sub

' === Einzelnes Element erzeugen
Function xml_add(xdoc, xparent, name, value)
    Dim xele ' Neues Element
    ' -- Unterelement erzeugen
    Set xele = xdoc.createElement(name)
    ' -- Wert setzen
    xele.text = value
    ' -- Element anfügen
    If xdoc.documentElement Is Nothing Then ' root-Element?
        Set xdoc.documentElement = xele ' Ja
    Else
        xparent.appendChild xele ' Nein
    End If
    Set xml_add = xele
End Function