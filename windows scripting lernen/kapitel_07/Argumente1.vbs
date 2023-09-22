' Argumente1.vbs
' Auslesen von Kommandozeilenargumenten
' =============================== 
Option Explicit
Dim i
Dim oArgumente 
Set oArgumente = WScript.Arguments
WScript.Echo("Es wurden " + _
    CStr(oArgumente.Length) + _
    " Argumente übergeben.")
For i = 0 To oArgumente.Count - 1
    WScript.Echo("Argument " + _
    CStr(i) + " = " + _
    oArgumente(i))
Next