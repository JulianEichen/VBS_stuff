' ------------------------------------------
' Skriptname: VerzeichnisstrukturAnlegen.vbs
' ------------------------------------------
' Dieses Skript legt im Dateisystem eine 
' Verzeichnisstruktur gemäß den Vorgaben einer
' XML-Datei an.
' ------------------------------------------
' Verwendet SCRRun, MSXML, WSH-Objekte
' ------------------------------------------
Option Explicit
' Deklaration der Variablen
Dim FSO
Dim XMLDocument
Dim WSHShell
Dim Eingabedatei
Dim StartKnoten
' Parameter
Const Basisverzeichnis = "T:\"
' Notwendige COM-Objekte erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
Set XMLDocument = CreateObject("Msxml2.DOMDocument")
XMLDocument.async = False
Set WSHShell = CreateObject("Wscript.shell")
' Basisverzeichnis erzeugen, wenn nicht vorhanden
If Not FSO.FolderExists(BasisVerzeichnis) Then
    WScript.Echo "Basisverzeichnis " & Basisverzeichnis & " wird erzeugt..."
    FSO.CreateFolder(BasisVerzeichnis)
End If
Eingabedatei = WSHShell.CurrentDirectory & "/Verzeichnisstruktur.xml"
' Lade die XML-Datei
WScript.echo "Lade " & Eingabedatei
XMLDocument.Load Eingabedatei
XMLDocument.async = False
' Rekursion starten mit Wurzelknoten
Set StartKnoten = XMLDocument.documentElement
' ALTERNATIV:
'Set StartKnoten = XMLDocument.SelectSingleNode("/VerzeichnisStruktur
/Verzeichnis[@Name='Websites']")
If Not StartKnoten Is Nothing Then
    VerzeichnisseAnlegen StartKnoten,Basisverzeichnis
Else
    WScript.Echo "Kein Startknoten!"
End If


' === Rekursive Hilfsroutine zum Anlegen der Verzeichnisse
Sub VerzeichnisseAnlegen(AktKnoten, AktVerz)
    Dim Unterknoten
    Dim NeuerName
    Dim NeuerPfad
    Dim i
    Dim Knoten
    Dim Ordner
    ' Schleife über alle Unterknoten
    Set Unterknoten = AktKnoten.childNodes
    For i = 0 To Unterknoten.length - 1
        Set Knoten = Unterknoten.Item(i)
        If Knoten.nodeType = 1 Then
            ' Knoten auslesen und neuen Verzeichnisnamen erzeugen
            NeuerName = Knoten.GetAttribute("Name")
            NeuerPfad = AktVerz & "\" & NeuerName
            If Not FSO.FolderExists(NeuerPfad) Then
                ' Verzeichnis erzeugen
                WScript.Echo "Verzeichnis " & NeuerPfad & " wird erzeugt..."
                Set Ordner = FSO.CreateFolder(NeuerPfad)
            Else
                WScript.Echo "Verzeichnis " & NeuerPfad & " ist bereits vorhanden!"
            End If
            ' Rekursion
            VerzeichnisseAnlegen Knoten, NeuerPfad
        End If
    Next
End Sub