Dim ProgramPath, ObjShell, ProgramArgs, WaitOnReturn,intWindowStyle,ScriptEngine
Set ObjShell = CreateObject("WScript.Shell")
Set objStdOut = WScript.StdOut

ProgramPath="C:\Users\Julian\Documents\VBS_stuff\VBS_stuff\cmd_scripts\helloworld.vbs"
ProgramArgs=""
intWindowStyle=1
WaitOnReturn=True
ScriptEngine="CScript.exe"

objStdOut.WriteLine "Running: " & ProgramPath
objStdOut.WriteBlankLines(1)
objStdOut.WriteLine "With" 
objStdOut.WriteLine "Args: " & ProgramArgs
objStdOut.WriteBlankLines(1)
objStdOut.WriteLine "ScriptEngine: " & ScriptEngine
objStdOut.WriteBlankLines(1)


Set Process=ObjShell.Exec (ScriptEngine & space (1) & Chr(34) & ProgramPath & Chr (34) & Space (1) & ProgramArgs)
Do While Process.Status=0
    'Currently Waiting on the program to finish execution.
    WScript.Sleep 300
Loop

objStdOut.WriteLine "Stdout: "
Set objStdOut = WScript.StdOut
objStdOut.WriteLine Process.StdOut.ReadAll 