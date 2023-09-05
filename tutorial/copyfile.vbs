Dim FSO, originalPath, targetPath
Set FSO = CreateObject("Scripting.FileSystemObject")

FSO.copyFile "Urpfad", "Zielpfad", True