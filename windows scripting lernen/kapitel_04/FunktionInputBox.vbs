' FunktionInputbox.vbs
' Die integrierte Funktion InputBox
' verwendet: keine weiteren Komponenten
' =============================== 
Dim Wert 
Wert = InputBox("Bitte geben Sie Ihren Namen ein", "Eingabe", "Oliver")
WScript.Echo("Hallo " + Wert + "!")
