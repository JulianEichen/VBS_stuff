' CSVDatei_Beschreiben.vbs
' Schreiben von Werten in eine CSV-Datei
' verwendet: SCRRun
' ===============================
Option Explicit
' Variablen deklarieren
Dim Benutzername, Vorname, Nachname, Geburtstag, Abteilungsnummer
Dim FSO,Datei,i
Const ForWriting = 2
Const ForAppending=8
' Arrays mit Werten füllen
Benutzername=Array("HugoHastig","WilliWinzig", _ 
    "StefanDerrick","HarryKlein")
Vorname=Array("Hugo","Willi","Stefan","Harry")
Nachname=Array("Hastig","Winzig","Derrick","Klein")
Geburtstag=Array("01.08.1935","27.05.1944","23.12.1912","14.02.1958")
Abteilungsnummer=Array("1","1","2","3")
'Erzeugen eines FSO-Objekts
Set FSO = CreateObject("Scripting.FileSystemObject")
'Erzeugen der Datei
Set Datei = fso.OpenTextFile("Benutzerliste.csv",ForAppending)
'Schreiben der einzelnen Werte
For i=0 to UBound(Benutzername)
    Datei.Write Benutzername(i) & ";"
    Datei.Write Vorname(i) & ";"
    Datei.Write Nachname(i) & ";"
    Datei.Write Geburtstag(i) & ";"
    Datei.Writeline Abteilungsnummer(i)
Next
Datei.Close