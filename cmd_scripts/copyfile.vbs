'Kopiert ein File zu einem neuen Ort
'erwartet 2 Pfade:
'   originalPath - Pfad zum File
'   targetPath - PFad zum Zielort

originalPath = WScript.Arguments.Item(0)
targetPath = WScript.Arguments.Item(1)

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.copyFile originalPath, targetPath, True

Set FSO = Nothing