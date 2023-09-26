' SetzeDateiAttribute.vbs
' Setzen von Dateieigenschaften
' verwendet: SCRRun
' ===============================
Option Explicit
' Deklaration der Variablen
Dim FSO, Datei, Attributwerte
' Konstanten definieren
Const Dateiname="C:\boot.ini"
'Objekt erzeugen
Set FSO = CreateObject("Scripting.FileSystemObject")
' Gibt es die Datei überhaupt?
If FSO.FileExists(Dateiname) Then
    ' Ja, also eine Verbindung herstellen
    Set Datei = FSO.GetFile(Dateiname)
    WScript.Echo "Größe der Datei: " & Datei.Size & " Bytes."
    WScript.Echo "Typ der Datei: " & Datei.Type
    WScript.Echo "Attribute der Datei: " & Datei.Attributes
    WScript.Echo "Erstellt am " & Datei.DateCreated
    WScript.Echo "Geändert am " & Datei.DateLastAccessed
    WScript.Echo "Letzter Zugriff " & Datei.DateLastModified
    Attributwerte=HoleAttribute(Datei)
    WScript.Echo "Die Datei " & Dateiname & " hat die Attributwerte
    [" & _ Attributwerte & "]"
    ' Entfernen der Attribute
    Datei.Attributes = Datei.Attributes and not 0
    Datei.Attributes = Datei.Attributes and not 1
    Datei.Attributes = Datei.Attributes and not 2
    Datei.Attributes = Datei.Attributes and not 4
    Datei.Attributes = Datei.Attributes and not 32
    Datei.Attributes = Datei.Attributes and not 64
    Datei.Attributes = Datei.Attributes and not 128
    Attributwerte=HoleAttribute(Datei)
    WScript.Echo "Die Datei " & Dateiname & " hat die Attributwerte
    [" & _ Attributwerte & "]"
    Attributwerte=""
    'Setzen der Attribute
    Datei.Attributes = Datei.Attributes or 0
    Datei.Attributes = Datei.Attributes or 1
    Datei.Attributes = Datei.Attributes or 2
    Datei.Attributes = Datei.Attributes or 4
    Datei.Attributes = Datei.Attributes or 32
    Datei.Attributes = Datei.Attributes or 64
    Datei.Attributes = Datei.Attributes or 128
    Attributwerte=HoleAttribute(Datei)
    WScript.Echo "Die Datei " & Dateiname & " hat die Attributwerte
    [" & _ Trim(Attributwerte) & "]"
Else
    WScript.Echo "Datei " & Dateiname & " nicht gefunden!"
End If
Private Function HoleAttribute(Handle)
' Hilfsroutine : Aufschlüsseln von Dateieigenschaften
' Deklaration der Variablen
Dim Ausgabe
' Normal-Flag gesetzt
If Datei.attributes and 0 Then Ausgabe = Ausgabe & "Normal "
' Nur-Lesen-Flag gesetzt
If Datei.attributes and 1 Then Ausgabe = Ausgabe & "Nur-Lesen "
' Versteckt-Flag gesetzt
If Datei.attributes and 2 Then Ausgabe = Ausgabe & "Versteckt "
' System-Flag gesetzt
If Datei.attributes and 4 Then Ausgabe = Ausgabe & "System "
' Archiv-Flag gesetzt
If Datei.attributes and 32 Then Ausgabe = Ausgabe & "Archiv "
' Alias-Flag gesetzt
If Datei.attributes and 64 Then Ausgabe = Ausgabe & "Link "
' Komprimiert-Flag gesetzt
If Datei.attributes and 128 Then Ausgabe = Ausgabe & "Komprimiert "
' Werte zurückgeben
HoleAttribute=Trim(Ausgabe)
End Function