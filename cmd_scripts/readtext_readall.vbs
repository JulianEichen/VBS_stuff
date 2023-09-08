Option Explicit

Dim fso, objStdOut, file, word

Set fso = CreateObject("Scripting.FileSystemObject")
Set objStdOut = WScript.StdOut

Set file = fso.OpenTextFile(WScript.Arguments.Item(0))

objStdOut.Write file.ReadAll

FOR EACH word IN file
    objStdOut.Write word
NEXT

file.close