' ============================================================
'  IP Geo Monitor - silent launcher (no console window)
'  Double-click this file to start the monitor.
' ============================================================
Option Explicit

Dim sh, fso, here, ps1, cmd
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

here = fso.GetParentFolderName(WScript.ScriptFullName)
ps1  = here & "\IpGeoMonitor.ps1"

If Not fso.FileExists(ps1) Then
    MsgBox "IpGeoMonitor.ps1 not found:" & vbCrLf & ps1, 16, "IP Geo Monitor"
    WScript.Quit 1
End If

sh.CurrentDirectory = here
cmd = "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & ps1 & """"
sh.Run cmd, 0, False
