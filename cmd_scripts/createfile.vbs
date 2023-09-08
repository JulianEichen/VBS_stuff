Option Explicit

Dim fso, objStdOut, targetPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set objStdout = WScript.StdOut

targetPath = fso.BuildPath(WScript.Arguments.Item(0), WScript.Arguments.Item(1))