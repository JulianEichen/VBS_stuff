' Konvertierungsfunktionen.vbs
' Konvertierungsfunktionen von VBScript
' verwendet: keine weiteren Komponenten
' =============================== 
On Error Resume Next
' Asc - Konvertierung eines Zeichens in einen ANSI-Wert
WScript.Echo "Asc-Konvertierung von 'A' = " & Asc("A")
WScript.Echo "Asc-Konvertierung von 'Abc' = " & Asc("Abc")
WScript.Echo "Asc-Konvertierung von 'dEF' = " & Asc("dEF")
' CBool - Konvertierung von Bedingungen 
WScript.Echo "CBool-Konvertierung von '5 = 5' = " & CBool(5 = 5)
WScript.Echo "CBool-Konvertierung von '0' = " & CBool(0)
WScript.Echo "CBool-Konvertierung von '1' = " & CBool(1)
WScript.Echo "CBool-Konvertierung von '4 = 5' = " & CBool(4 = 5)
' CByte - Konvertierung einer Zeichenkette in einen Byte-Wert
WScript.Echo "CByte-Konvertierung von '123' = " & CByte("123")
WScript.Echo "CByte-Konvertierung von 'ABC' = " & CByte("ABC")
If Err.Number <> 0 Then 
    WScript.Echo "Es ist ein Fehler bei dieser Konvertierung aufgetreten"
    Err.Clear()
End If
' CCur - Konvertierung einer Zeichenkette in einen Währungsbetrag
WScript.Echo "CCur-Konvertierung von '12,34' = " & CCur("12,34")
WScript.Echo "CCur-Konvertierung von '4.321,99' = " & CCur("4.321,99")
' CDate - Konvertierung einer Zeichenkette in einen Datumswert
WScript.Echo "CDate-Konvertierung von '11.4.1975' = " & CDate("11.4.1975")
WScript.Echo "CDate-Konvertierung von '31.7.02' = " & CDate("31.7.02")
' CDbl - Konvertierung einer Zeichenkette in einen Double-Wert
WScript.Echo "CDbl-Konvertierung von '55,43' = " & CDbl("55,43")
WScript.Echo "CDbl-Konvertierung von '2312,32323' = " & CDbl("2312,32323")
' Chr - Konvertierung einer Zahl in das entsprechende ASCII-Zeichen
WScript.Echo "Chr-Konvertierung von '65' = " & Chr(65)
WScript.Echo "Chr-Konvertierung von '123' = " & Chr(123)
' CInt - Konvertierung einer Zeichenkette in einen Double-Wert
WScript.Echo "CInt-Konvertierung von '65' = " & CInt("65")
WScript.Echo "CInt-Konvertierung von '123,2323' = " & CInt("123,2323")
WScript.Echo "CInt-Konvertierung von '12a' = " & CInt("12a2")
If Err.Number <> 0 Then 
    WScript.Echo "Es ist ein Fehler bei dieser Konvertierung aufgetreten"
    Err.Clear()
End If
' CLng - Konvertierung einer Zeichenkette in einen Long-Wert
WScript.Echo "CLng-Konvertierung von '65' = " & CLng("65")
WScript.Echo "CLng-Konvertierung von '123' = " & CLng("123")
' CSng - Konvertierung einer Zeichenkette in einen Single-Wert
WScript.Echo "CSng-Konvertierung von '65' = " & CSng("65")
WScript.Echo "CSng-Konvertierung von '123' = " & CSng("123")
' CStr - Konvertierung einer Zeichenkette in einen String-Wert
WScript.Echo "CStr-Konvertierung von '65' = " & CStr("65")
WScript.Echo "CStr-Konvertierung von '123' = " & CStr("123")
' Hex - Konvertierung einer Zeichenkette in einen Hex-Wert
WScript.Echo "Hex-Konvertierung von '65' = " & Hex("65")
WScript.Echo "Hex-Konvertierung von '123' = " & Hex("123")
' Oct - Konvertierung einer Zeichenkette in einen Oktal-Wert
WScript.Echo "Oct-Konvertierung von '65' = " & Oct("65")
WScript.Echo "Oct-Konvertierung von '123' = " & Oct("123")