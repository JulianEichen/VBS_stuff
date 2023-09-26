' DatenbankTabelle_Lesen.vbs
' Lesen der Benutzerliste aus einer Access-Datenbank
' verwendet: ADO
' ===================================================
'Deklarieren der Variablen
Dim DBConnection, SqlString, Ergebnismenge
'Definieren der Konstanten
Const Verbindung="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=.\User.MDB;"
' Erstellen eines Connection-Objekts
Set DBConnection = CreateObject("ADODB.Connection")
' Öffnen der Verbindung zur Datenbank
' Die MDB-Datei muss im selben Verzeichnis liegen wie das Skript
DBConnection.Open Verbindung
' Abfrage der Tabelle Benutzer
SqlString="SELECT * FROM Benutzer"
' Ausführen der Abfrage und Rückgabe eines Recordsets
Set Ergebnismenge = DBConnection.Execute(SqlString)
' An den Anfang des Recordsets springen
Ergebnismenge.MoveFirst
' Durchlaufen des gesamten Ergebnisses
Do While Not Ergebnismenge.eof
    ' Ausgabe der Daten
    WScript.echo "Ausgabe der Daten für " & Ergebnismenge("Benutzername")
    WScript.Echo Ergebnismenge("Vorname")
    WScript.Echo Ergebnismenge("Nachname")
    WScript.Echo Ergebnismenge("Geburtsdatum")
    WScript.Echo Ergebnismenge("Abteilungsnummer")
    WScript.Echo "################"
    ' Datensatzzeiger auf den nächsten Datensatz positionieren
    Ergebnismenge.MoveNext
Loop
' Recordset schließen
Ergebnismenge.Close
' Verbindung schließen
DBConnection.Close