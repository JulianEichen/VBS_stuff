' FunktionRnd.vbs
' Erstellen von Zufallszahlen
' verwendet: keine weiteren Komponenten
' =============================== 
Dim Untergrenze, Obergrenze, Zufallszahl
Randomize()
Untergrenze = 0
Obergrenze = 100
Zufallszahl = Int((Obergrenze - Untergrenze + 1) * Rnd() + Untergrenze)
WScript.Echo("Eine Zufallszahl zwischen " & CStr(Untergrenze) & _
 " und " & CStr(Obergrenze) & " = " & CStr(Zufallszahl))
 
Untergrenze = 10
Obergrenze = 20
Zufallszahl = Int((Obergrenze - Untergrenze + 1) * Rnd() + Untergrenze)
WScript.Echo("Eine Zufallszahl zwischen " & CStr(Untergrenze) & _
 " und " & CStr(Obergrenze) & " = " & CStr(Zufallszahl))