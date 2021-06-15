Dim WinScriptHost

Set oAL = CreateObject("System.Collections.ArrayList")
For Each oItem In Wscript.Arguments: oAL.Add oItem: Next

path = """powershell.exe"""
param = " -EP Bypass -File goodTimes.ps1 " & Join(oAL.ToArray, " ")

Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.Run path & param, 0
Set WinScriptHost = Nothing
