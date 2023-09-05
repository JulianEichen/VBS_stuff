Option Explicit
Dim Mode
Mode = InputBox("Bitte wählen Sie die Rechenart aus. Einfach +, -, * oder / eingeben.", "Modus", "")
If Mode = "" Then 
    Wscript.Quit
End If
Dim a,b, Result
a = InputBox("Bitte Wer 1 Eingeben", "Eingabebox", "Hier die Zahl eingeben")
b = InputBox("Bitte Wer 1 Eingeben", "Eingabebox", "Hier die Zahl eingeben")
If Mode = "+" Then
    Result = a -- b
ElseIf
