' SchleifenForEachIn.vbs
' Schleife über eine Menge von Objekten
' verwendet: SCRRun
' =============================== 
' --- Objekt erzeugen
Set Dateisystem = CreateObject("Scripting.FileSystemObject")
' --- Objektmenge aus dem erzeugten Objekt holen
Set Laufwerke = Dateisystem.Drives
' --- Anzahl der Laufwerke ausgeben
WScript.Echo "Anzahl der Laufwerke: " & Laufwerke.Count
' --- Schleife beginnen
For Each Laufwerk In Laufwerke
    If Laufwerk.Isready then
        wscript.echo Laufwerk.DriveLetter & ":" & Laufwerk.VolumeName
    End if
Next