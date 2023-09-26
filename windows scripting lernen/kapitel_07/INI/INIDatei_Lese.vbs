' INIDatei_Lesen.vbs
' Lesen eines bestimmten Eintrags aus einer INI-Datei
' Verwendet: SCRRun
' ===============================
Option Explicit
' Variablen deklarieren
Dim FSO, IniDatei, Zeile, EintragWert, Zeichen
Dim EintragGefunden, SektionGefunden
Dim I
' Konstanten definieren
Const ForReading = 1
Const Dateiname="C:\boot.INI"
Const Sektion="boot loader"
Const Eintrag="default"
' Erzeugen eines FSO-Objektes
Set FSO = CreateObject("Scripting.FileSystemObject")
' INI-Datei zum Lesen öffnen
Set IniDatei = FSO.OpenTextFile(Dateiname, ForReading)
' Durchlaufe die gesamte Datei
Do While Not IniDatei.AtEndOfStream
    Zeile=IniDatei.readline
    ' Wenn die aktuelle Zeile eine Sektion kennzeichnet
    If Left(Zeile,1)="[" Then
        ' Ist es die gesuchte Sektion?
        If UCase(Mid(Zeile,2,Len(Sektion)))=UCase(Sektion) Then
            ' Ja
            SektionGefunden=True
        Else
            ' Nein
            SektionGefunden=False
        End If
    Else
        If SektionGefunden Then
            'Ist die aktuelle Zeile der gesuchte Eintrag?
            If UCase(Left(Zeile,Len(Eintrag)))=UCase(Eintrag) Then
                I = Len(Eintrag)+1
                ' So lange wiederholen, bis der Eintrag gefunden wurde
                Do While I<Len(Zeile)
                    Zeichen = Mid(Zeile,I,1)
                    ' Suche das Gleichheitszeichen
                    If Zeichen="=" Then
                        ' Ermitteln des Wertes
                        EintragWert=Right(Zeile,Len(Zeile)-I)
                        ' Abbruchbedingung setzen
                        I=Len(Zeile)
                        EintragGefunden=True
                    Else
                        I=I+1
                    End If
                Loop
            End If
        End If
    End If
Loop
' Datei schließen
IniDatei.close
' Wert zurückgeben
WScript.Echo "Der gesuchte Wert ist: " & Trim(EintragWert)
