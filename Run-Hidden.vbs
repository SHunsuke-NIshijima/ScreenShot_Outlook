Set objShell = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & "\Send-Screenshot-Email.ps1""", 0, False
