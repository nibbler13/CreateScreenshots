#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=D:\Scripts\create_screenshots\icon.ico
;~ #AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 0.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для создания разовых скриншотов)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555)
#pragma compile(ProductName, create_screenshot)

AutoItSetOption("TrayAutoPause", 0)

#include <ScreenCapture.au3>
#include <WinAPIShPath.au3>

Local $mainPath = ""
Local $current_pc_name = @ComputerName
Local $aCmdLine = _WinAPI_CommandLineToArgv($CmdLineRaw)

If UBound($aCmdLine) > 1 Then $mainPath = $aCmdLine[1]

If UBound($aCmdLine) <= 1 Or $mainPath = "/?" Or $mainPath = "-h" Or $mainPath = "--help" Then
	ConsoleWrite("Usage:" & @CRLF & "1st parameter - main path to saving" & @CRLF & _
		"Example: " & @ScriptName & " \\mscs-fs-01\InfomonitorsScreenshots\")
	Exit -1
EndIf

;~ $mainPath = "C:\"

If Not FileExists($mainPath) Then Exit -2

_ScreenCapture_SetJPGQuality(25)
Local $screenShot = _ScreenCapture_Capture("", 0, 0, -1, -1)
Exit _ScreenCapture_SaveImage($mainPath & @ComputerName & ".jpg", $screenShot, True)