Option Explicit

Dim fso, objStdOut, file, word

Set fso = CreateObject("Scripting.FileSystemObject")
Set objStdOut = WScript.StdOut

Set file = fso.OpenTextFile(WScript.Arguments.Item(0))

objStdOut.Write file.ReadAll

file.close