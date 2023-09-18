Set control = CreateObject("MSScriptControl.ScriptControl")
control.Language = "vbscript"
Set objOL = CreateObject("Wscript.Shell")
Set WshShellExec = WshShell.exec("whoami")
MsgBox (WshShellExec.StdOut.ReadAll)
