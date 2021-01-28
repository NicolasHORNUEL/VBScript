'L'application Serveur PCC a besoin du lecteur réseau Z:\\Master-xmf\Komo cip4

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject("WScript.Shell")

If oFSO.GetDrive("Z").IsReady = True Then
	Set ProcessList = GetObject("Winmgmts:").ExecQuery("Select * from Win32_Process")
	For Each objProcess in ProcessList
		myList = myList & vbCr & objProcess.Name
	Next
	If InStr(myList,"PCC.exe") = 0 Then
		oShell.Run ("Z:")
		oShell.Run """C:\Program Files\Komori\PCC2.29\PCC.exe"""
		WScript.Sleep 2000
		oShell.AppActivate "PCC"
		WScript.Sleep 100
		oShell.SendKeys ("^{a}")
	End If
else
	oShell.Run "taskkill /im PCC.exe", 1, True
End if


