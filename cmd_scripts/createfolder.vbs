Option Explicit

Dim fso, targetPath, objStdOut
  
Set fso = CreateObject("Scripting.FileSystemObject")
Set objStdout = WScript.StdOut

targetPath = fso.BuildPath(WScript.Arguments.Item(0), WScript.Arguments.Item(1))

If NOT fso.folderExists(targetPath) Then
    fso.CreateFolder (targetPath)
    objStdOut.Write "Folder Created"
Else
    objStdOut.Write "Folder Already Exists"
End If