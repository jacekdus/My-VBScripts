'/// CONFIG

Set WshShell = WScript.CreateObject("WScript.Shell")
Dim userName
Dim mainPath

userName = "user"
mainPath = "C:\Users\" & userName & "\Desktop\"

Sub runOnErrorResume(exeFile)
	On Error Resume Next	
		Do
			Err.Clear
			statusCode = WshShell.Run (exeFile, 1, true)
			'msgbox(statusCode)
			WScript.Sleep(5000)
		Loop While ( Err.Number <> 0)
	On Error GoTo 0
End Sub

'/// MAIN

runOnErrorResume(mainPath & "Programy\LoadVPN.bat")	'* Run VPN
runOnErrorResume(mainPath & "Programy\Log.lnk")		'* Run disks

'* Run anything else that needs disks first:
WshShell.Run("""C:\Program Files\Belkin\Network USB Hub Control Center\Connect.exe""")
WshShell.Run("""C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Outlook 2013.lnk""")
WshShell.Run("excel " & mainPath & "GODZINY.lnk")
WshShell.Run(mainPath & "Programy\skrypt.vbs")
