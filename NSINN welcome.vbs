dim speechobject
public strUser
set speechobject=createobject("sapi.spvoice")
strUser = CreateObject("WScript.Network").UserName
speechobject.volume = 25
speechobject.speak "Hello " & strUser & "!"