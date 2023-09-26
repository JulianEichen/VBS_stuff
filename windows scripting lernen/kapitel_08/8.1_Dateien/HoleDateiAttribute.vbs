' HoleDateiAttribute.vbs
' Ermitteln von Dateieigenschaften
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, Datei, Ausgabe
' Konstanten definieren
Const Dateiname="WSL_Kapitel08.doc"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Gibt es die Datei überhaupt?
if FSO.FileExists(Dateiname) then
    ' Ja, also eine Verbindung herstellen
    Set Datei = FSO.GetFile(Dateiname)
    WScript.Echo "Größe der Datei: " & Datei.Size & " Bytes."
    WScript.Echo "Typ der Datei: " & Datei.Type
    WScript.Echo "Attribute der Datei: " & Datei.Attributes
    WScript.Echo "Erstellt am " & Datei.DateCreated
    WScript.Echo "Geändert am " & Datei.DateLastAccessed
    WScript.Echo "Letzter Zugriff " & Datei.DateLastModified
    ' Dateiattribute ermitteln
    If Datei.attributes and 0 Then Ausgabe = Ausgabe & "Normal "
    If Datei.attributes and 1 Then Ausgabe = Ausgabe & "Nur-Lesen "
    If Datei.attributes and 2 Then Ausgabe = Ausgabe & "Versteckt "
    If Datei.attributes and 4 Then Ausgabe = Ausgabe & "System "
    If Datei.attributes and 32 Then Ausgabe = Ausgabe & "Archiv "
    If Datei.attributes and 64 Then Ausgabe = Ausgabe & "Link "
    If Datei.attributes and 128 Then Ausgabe = Ausgabe & "Komprimiert "
else
    WScript.Echo "Datei " & Dateiname & " nicht gefunden!"
end if
WScript.Echo "Die Datei " & Dateiname & _
" hat die Attributwerte [" & Trim(Ausgabe) & "]"