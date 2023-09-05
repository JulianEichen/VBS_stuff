'Loescht ein File
' Erwartet 1 Pfad
'   filePath - zu loeschendes File

filePath = WScript.Arguments.Item(0)

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

FSO.DeleteFile filePath, False