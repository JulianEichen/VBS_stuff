' Argumente2.vbs
' Auslesen von Kommandozeilenargumenten
' verwendet: WSHRun
' Übergabe unterschiedlicher Argumente 
' Dabei muss jedes Argument mit einer Kennzeichnung
' übergeben werden.
' =============================== 
Option Explicit
Dim i
Dim oArgumente
Dim sDatei 
Dim sVerzeichnis
Dim sComputer
Set oArgumente = WScript.Arguments
' Bei Argument 0 beginnen 
i = 0 
If oArgumente.Count > 0 Then
 Do
 If UCase(oArgumente(i)) = "-V" Or _
 UCase(oArgumente(i)) = "-VERZEICHNIS" Then
 ' --- Verzeichnisargument
 i = i + 1
 sVerzeichnis = oArgumente(i)
 
 Elseif UCase(oArgumente(i)) = "-D" Or _
 UCase(oArgumente(i)) = "-DATEI" Then
 ' --- Dateiargument
 i = i + 1
 sDatei = oArgumente(i)
 
 Elseif UCase(oArgumente(i)) = "/C" Or _
 UCase(oArgumente(i)) = "/COMPUTER" Then
 i = i + 1
 sComputer = oArgumente(i)
 
 End If
 i = i + 1
 Loop Until i>=oArgumente.Count
End if
If sDatei = "" And sVerzeichnis = "" And sComputer = "" Then
    ' Es wurde kein Argument übergeben
    WScript.Echo("Es wurden keine oder falsche Argumente übergeben.")
    WScript.Echo(vbTab + _
    "-d dateiname oder -datei dateiname")
    WScript.Echo(vbTab + _
    "-v verzeichnisname oder -verzeichnis verzeichnisname")
    WScript.Echo(vbTab + _
    "-c computername oder -computer computername")
 
Else
    If sDatei <> "" Then
    WScript.Echo("Datei = " + sDatei)
    ' Aktionen für die übergebene Datei
    End If

    If sVerzeichnis <> "" Then
         WScript.Echo("Datei = " + sVerzeichnis)
        ' Aktionen für das übergebene Verzeichnis
    End If
 
    If sComputer <> "" Then
        WScript.Echo("Computer = " + sComputer)
        ' Aktionen für den übergebenen Computer
    End If
 
End If