' wsh-versionsnummer.vbs
' Ausgabe der Versionsnummern des WSH und von VBScript
' verwendete Komponenten: WSH, VBS
' ================================

WScript.Echo _
"Dies ist der " & WScript.Name & _
" Version " &WScript.Version

WScript.Echo _
"Die ist die Sprache " & ScriptEngine & _
" Version " & _
ScriptEngineMajorVersion & "." & _
ScriptEngineMinorVersion & "." & _
ScriptEngineBuildVersion