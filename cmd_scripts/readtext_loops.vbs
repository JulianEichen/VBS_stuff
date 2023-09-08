Option Explicit

Dim dict, file, fso, line, objStdOut, row 

Set fso = CreateObject("Scripting.FilesyStemObject")
Set objStdOut = WScript.StdOut

Set file = fso.OpenTextFile(WScript.Arguments.Item(0))
Set dict = CreateObject("Scripting.Dictionary")

row = 0
Do Until file.AtEndOfStream
  line = file.Readline
  dict.Add row, line
  row = row + 1
Loop

file.close

FOR EACH line IN dict.Items
    objStdOut.Write line
    objStdOut.WriteBlankLines(1)
NEXT