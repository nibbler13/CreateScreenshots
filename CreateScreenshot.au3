#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=D:\Scripts\create_screenshots\icon.ico
;~ #AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 0.2)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для создания скриншотов)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555)
#pragma compile(ProductName, create_screenshot)

AutoItSetOption("TrayAutoPause", 0)
;~ AutoItSetOption("TrayIconDebug", 1)

#include <ScreenCapture.au3>
#include <WinAPIShPath.au3>
#include <Array.au3>
#include <Date.au3>
#include <File.au3>

Local $delay = 60 * 1000
Local $mainPath = ""
Local $optionalPath = ""
Local $current_pc_name = @ComputerName
Local $tempFolder = StringSplit(@SystemDir, "\")[1] & "\Temp\"
Local $errStr = "===ERROR=== "

If Not FileExists($tempFolder) Then
   If Not DirCreate($tempFolder) Then
	  $error = True
	  ConsoleWrite($errStr & "Cannot create folder " & $tempFolder & @CRLF)
   EndIf
EndIf

Local $logFilePath = $tempFolder & @ScriptName & ".log"
Local $logFile = FileOpen($logFilePath, $FO_OVERWRITE)
Local $aCmdLine = _WinAPI_CommandLineToArgv($CmdLineRaw)

If UBound($aCmdLine) > 1 Then
	If StringInStr($aCmdLine[1], "|") Then
		Local $tmpStr = StringSplit($aCmdLine[1], "|", $STR_NOCOUNT)
		$mainPath = $tmpStr[0]
		$optionalPath = $tmpStr[1]
	Else
		$mainPath = $aCmdLine[1]
	EndIf

	If UBound($aCmdLine) > 2 Then $optionalPath = $aCmdLine[2]
EndIf

If UBound($aCmdLine) <= 1 Or $mainPath = "/?" Or $mainPath = "-h" Or $mainPath = "--help" Then
	ConsoleWrite("Usage:" & @CRLF & "1st parameter - main path to saving" & @CRLF & "2nd parameter - optional path to saving" & @CRLF & _
		"Example: " & @ScriptName & " \\mscs-fs-01\InfomonitorsScreenshots\ \\nnkk-fs\InfomonitorsScreenshots\")
	Exit
EndIf

;~ $mainPath = "C:\"
;~ $optionalPath = "D:\"

ToLog("Main path: " & $mainPath)
ToLog("Optional path: " & $optionalPath)

If Not FileExists($mainPath) Then
	ToLog($errStr & "Main path doesn't exists")
EndIf

If $optionalPath <> "" And Not FileExists($optionalPath) Then ToLog("Optional path doesn't exists")

Local $organizationName = StringSplit($current_pc_name, "-", $STR_NOCOUNT)[0]
Local $lastTime = @HOUR
Sleep($delay)
_ScreenCapture_SetJPGQuality(25)

While True
	Local $currentTime = @HOUR

	If $currentTime <> $lastTime Then
		Local $screenShot = _ScreenCapture_Capture("", 0, 0, -1, -1)
		WriteFileTo($mainPath, $screenShot)
		If $optionalPath <> "" Then WriteFileTo($optionalPath, $screenShot)
		$lastTime = $currentTime
		_WinAPI_DeleteObject($screenShot)
	EndIf

	Sleep($delay)
WEnd

Func WriteFileTo($path, $data)
	ToLog("Trying to write screenshot to: " & $path)
	Local $currentDate = @YEAR & @MON & @MDAY

	$path &= $organizationName
	If Not CheckFolder($path) Then Return

	Local $folderList = _FileListToArray($path, "*", $FLTA_FOLDERS)

	If IsArray($folderList) Then
		For $i = 1 To $folderList[0]
			If $folderList[$i] <> $currentDate Then
				ToLog("Removing directory: " & $path & "\" & $folderList[$i] & " - " & DirRemove($path & "\" & $folderList[$i], $DIR_REMOVE))
			EndIf
		Next
	EndIf

	$path &= "\" & $currentDate
	If Not CheckFolder($path) Then Return

	$path &= "\" & @HOUR
	If Not CheckFolder($path) Then Return

	If Not _ScreenCapture_SaveImage($path & "\" & $current_pc_name & ".jpg", $screenShot, False) Then
		ToLog($errStr & "Cannot write the screenshot")
	Else
		ToLog("The screenshot succesfully saved to: " & $path & "\" & $current_pc_name & ".jpg")
	EndIf

    _GDIPlus_Startup()
    Local $bitmap = _GDIPlus_BitmapCreateFromHBITMAP($data)
	Local $scale = 200 / @DesktopWidth
    Local $scaled = _GDIPlus_ImageResize($bitmap, @DesktopWidth * $scale, @DesktopHeight * $scale, 7)
    Local $sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
    Local $tParams = _GDIPlus_ParamInit(1)
    Local $tData = DllStructCreate("int Quality")
    DllStructSetData($tData, "Quality", 60)
    Local $pData = DllStructGetPtr($tData)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, $pData)
    Local $pParams = DllStructGetPtr($tParams)

    If Not _GDIPlus_ImageSaveToFileEx($scaled, $path & "\" & "_preview_" & $current_pc_name & ".jpg", $sCLSID, $pParams) Then
		ToLog($errStr & "Cannot write the screenshot preview")
	Else
		ToLog("The screenshot succesfully saved to: " & $path & "\" & "_preview_" & $current_pc_name & ".jpg")
	EndIf

    _GDIPlus_BitmapDispose($bitmap)
    _GDIPlus_BitmapDispose($scaled)
    _GDIPlus_Shutdown()
EndFunc

Func CheckFolder($path)
	If Not FileExists($path) Then
		If Not DirCreate($path) Then
			ToLog($errStr & "Cannot create folder: " & $path)
			Return False
		EndIf
	EndIf

	Return True
EndFunc

Func ToLog($message)
   $message &= @CRLF
   ConsoleWrite(_NowCalc() & ": " & $message)
   _FileWriteLog($logFile, $message)
EndFunc