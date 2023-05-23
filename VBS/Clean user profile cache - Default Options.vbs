
Dim objShell
Set objShell=CreateObject("WScript.Shell")

'enter the PowerShell expression
strExpression=".\CacheClean.ps1 -Clean"

strCMD="C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -nologo -command " & Chr(34) &_
"&{" & strExpression &"}" & Chr(34)

'Uncomment next line for debugging
'WScript.Echo strCMD

'use 0 to hide window
objShell.Run strCMD,0









