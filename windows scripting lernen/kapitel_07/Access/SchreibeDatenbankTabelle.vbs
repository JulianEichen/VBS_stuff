' SchreibeDatenbankTabelle.vbs
' Schreiben von Werten in eine Datenbanktabelle
' verwendet: SCRRun, ADO
' ===========================================================================
' Variablen deklarieren
Dim DBConnection
Dim Tabelle
Dim FSO
Dim Datei,TextZeile,Ausgabe, Zaehler
' Konstanten für Datenzugriffe definieren
Const Verbindung="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=.\User.MDB;"
Const adOpenDynamic = 2
Const adLockOptimistic = 3
Const ForReading = 1
Const adOpenKeyset = 1
' FSO erstellen
Set FSO=CreateObject("Scripting.FileSystemObject")
' Datei zum Einlesen öffnen
Set Datei = FSO.OpenTextFile("Benutzerliste.csv", ForReading, False)
' Connection-Objekt erstellen
Set DBConnection = CreateObject("ADODB.Connection")
' Connection öffnen
' Die MDB-Datei muss im selben Verzeichnis liegen wie das Skript
DBConnection.Open Verbindung
' Recordset-Objekt erstellen
Set Tabelle = CreateObject("ADODB.Recordset")' verwendete Connection festlegen
Tabelle.ActiveConnection = DBConnection
' Zugriffsart festlegen
Tabelle.CursorType = adOpenDynamic
' Sperrart festlegen
Tabelle.LockType = adLockOptimistic' verwendete Quelle angeben
Tabelle.Source="Benutzer2"
' Tabelle öffnen
Tabelle.Open
Zaehler=0
' Gesamte Textdatei durchlaufen
While not Datei.AtEndOfStream
    ' Erste Zeile überlesen, enthält die Feldnamen
    if Zaehler=0 then TextZeile=Datei.Readline()
    ' Zeilenweise einlesen
    TextZeile=Datei.Readline()
    ' Werte trennen
    Benutzer=Split(TextZeile,";")
    ' Hinzufügen einer neuen Tabellenzeile
    Tabelle.AddNew
    ' Spalten mit Werten besetzen
    Tabelle("Benutzername") = Benutzer(0)
    Tabelle("Vorname") = Benutzer(1)
    Tabelle("Nachname") = Benutzer(2)
    Tabelle("Geburtsdatum") = CDate(Benutzer(3))
    Tabelle("Abteilungsnummer") = CInt(Benutzer(4))
    ' Änderungen schreiben
    Tabelle.Update
    Zaehler=Zaehler+1
Wend
'Objekte schließen
Tabelle.Close
DBConnection.Close
Datei.Close