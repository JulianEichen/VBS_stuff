' Objektmengenzugriff.vbs
' Zugriff auf einzelne Elemente einer Objektmenge
' verwendet: SCRRun
' ================================================================ 
' --- Objekt erzeugen
Set Dateisystem = CreateObject("Scripting.FileSystemObject")
' --- Objektmenge aus dem erzeugten Objekt holen
Set Laufwerke = Dateisystem.Drives
' --- Objekt einzeln ansprechen (ALTERNATIVE 1)
WScript.echo Laufwerke.Item("D:").VolumeName
' --- Objekt einzeln ansprechen (ALTERNATIVE 2)
WScript.echo Laufwerke("D:").VolumeName
' --- Objekt einzeln ansprechen (ALTERNATIVE 3)
WScript.echo Dateisystem.GetDrive("D:").Volumename