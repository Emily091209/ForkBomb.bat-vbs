Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "Enter file destination" & Chr(34), 0
Set WshShell = Nothing
