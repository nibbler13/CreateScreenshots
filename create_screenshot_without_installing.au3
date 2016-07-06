Local $fileName = "Screenshot.jpg"
Local $ip = "172.16.166."
For $i = 150 To 197
	Local $pcName = $ip & $i
   Local $command = '/c ' & @ScriptDir & '\psexec.exe \\' & $pcName & ' -u "budzdorov\" -p "" -i -c ' & @ScriptDir & _
	  '\nircmd.exe savescreenshot %SYSTEMDRIVE%\Temp\' & $fileName
	ConsoleWrite($command & @CRLF)
	ConsoleWrite(RunWait("cmd.exe " & $command) & @CRLF);, Default, @SW_HIDE) & @CRLF)
	Local $date = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC
	FileCopy("\\" & $pcName & "\C$\Temp\" & $fileName, @ScriptDir & "\" & $pcName & "_" & $date & "_" & $fileName)
Next