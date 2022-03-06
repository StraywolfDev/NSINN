Public WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")

public speechobject

set speechobject=createobject("sapi.spvoice")

speechobject.volume = 25

Public timeseconds

timeseconds = UserInput("Time? (In Seconds!) DO NOT ENTER 0")

Function UserInput( myPrompt )

    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
        
        WScript.StdOut.Write myPrompt & " "
        UserInput = WScript.StdIn.ReadLine
    Else
        
        UserInput = InputBox( myPrompt,"NSINN" )
    End If
End Function

Public red
red = timeseconds-1

speechobject.speak "T Minus " & red+1

WScript.Sleep 1000

Do Until red=0

speechobject.speak red
WScript.Sleep 1000
red=red-1

Loop

speechobject.speak "Self Destructing!"