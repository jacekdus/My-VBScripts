'/// CONFIG

Set WshShell = WScript.CreateObject("WScript.Shell")
Dim userName
Dim path

Dim exeName
Dim statusCode

userName = "user"
appPath = "C:\Users\" & userName & "\Desktop\"

Sub run(exeName)
	statusCode = WshShell.Run (exeName, 1, true)
End Sub

'/// MAIN

run(appPath & "loadVPN.bat")	'* Run VPN

On Error Resume Next			'* Run disks
	Do
		Err.Clear
		run(appPath & "log.lnk") 	
		WScript.Sleep(5000)
	Loop While ( Err.Number <> 0)
On Error GoTo 0

'* Run anything else that needs disks first:
WshShell.Run("""C:\Program Files\Belkin\Network USB Hub Control Center\Connect.exe""")
WshShell.Run("""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Outlook 2013.lnk""")
WshShell.Run(appPath & "Programy\skrypt.vbs")
