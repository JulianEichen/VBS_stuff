' INIDatei_Schreiben.vbs
' Schreiben in eine INI-Datei
' Verwendet: SCRRun
' ===============================
Option Explicit
' Variablen deklarieren
Dim FSO, IniDatei, TempDatei
Dim temp, TempFolder
Dim Zeile,AlterEintrag
Dim SektionGefunden
'Konstanten deklarieren
Const ForReading = 1
Const ForWriting = 2
Const Dateiname="g:\test\ITVisions.ini"
Const Eintrag="Website"
Const Inhalt="www.IT-Visions.de"
Const Sektion="Internet"
' -- Werte vorbelegen
AlterEintrag=False
SektionGefunden=False
' Erzeugen des FSO-Objektes
Set FSO = CreateObject("Scripting.FileSystemObject")
' INI-Datei öffnen
Set IniDatei = FSO.OpenTextFile(Dateiname, ForReading)
' Temporären Dateinamen erzeugen
temp = FSO.GetTempName
' Ermitteln des Temp-Verzeichnisses
TempFolder=FSO.GetSpecialFolder(2)
' Erzeugen der temporären Datei und gleichzeitiges Öffnen
Set TempDatei = FSO.CreateTextFile(TempFolder & "\" & temp, True)
' Solange das Ende der Datei nicht erreicht ist
Do While Not IniDatei.AtEndOfStream
    ' Zeilenweise einlesen
    Zeile=IniDatei.Readline
    ' Sektionsbeginn suchen
    If Left(Zeile,1)="[" Then
        If InStr(UCase(Zeile), "[" + UCase(Sektion) + "]") > 0 Then
            ' Ja
            SektionGefunden=True
        Else
            ' Nein
            SektionGefunden=False
        End If
    Else
        If SektionGefunden Then
            If UCase(Left(Zeile,Len(Eintrag)))=UCase(Eintrag) Then
                ' Eintrag gefunden
                ' Neuen Eintrag in Variable schreiben
                Zeile=Eintrag & "=" & Inhalt
                ' Merker setzen
                AlterEintrag=True
            End If
        End If
    End If
    'Zeile in temp. Datei schreiben
    TempDatei.Writeline(Zeile)
Loop
' Wenn es ein neuer Eintrag ist, dann Sektion und Eintrag schreiben
If Not AlterEintrag Then
    If Sektion <> "" Then TempDatei.Writeline("[" + Sektion + "]")
        TempDatei.Writeline(Eintrag + " =" + Inhalt)
End If
' Temporäre Datei schließen
TempDatei.Close
' INI-Datei schließen
IniDatei.Close
' Temporäre Datei kopieren und alte Datei überschreiben
FSO.CopyFile TempFolder & "\" & temp,Dateiname,True
' Temporäre Datei löschen
FSO.DeleteFile TempFolder & "\" & temp, True