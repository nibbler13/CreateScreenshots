#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 0.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для отображения скриншотов инфомониторов)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555)
#pragma compile(ProductName, create_print_statistics_for_all)

#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>
#include <ColorConstants.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <AD.au3>
#include <GUIScrollbars_Ex.au3>
#include <GDIPlus.au3>
;~ #include <File.au3>
;~ #include <FileConstants.au3>
;~ #include <Excel.au3>
;~ #include <GuiEdit.au3>
;~ #include <GuiScrollBars.au3>
;~ #include <EditConstants.au3>
;~ #include <Excel.au3>
;~ #include <String.au3>


#Region ==========================	 Variables	 ==========================
Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $iniFile = "\\budzdorov.ru\NETLOGON\Restart_Infoscreen\autorun_for_all_infomonitors_settings.ini"

If Not FileExists($iniFile) Then
	MsgBox($MB_ICONERROR, @ScriptName, "Cannot find the settings file: " & $iniFile)
	Exit
EndIf

Local $allSections = IniReadSectionNames($iniFile)
If @error Then
	MsgBox($MB_ICONERROR, @ScriptName, "Cannot find any sections in the settings file: " & $iniFile)
	Exit
EndIf

_ArrayDelete($allSections, 0)

Local $index
$index = _ArraySearch($allSections, "general")
If Not @error Then _ArrayDelete($allSections, $index)
$index = _ArraySearch($allSections, "mail")
If Not @error Then _ArrayDelete($allSections, $index)

If Not UBound($allSections) Then
	MsgBox($MB_ICONERROR, @ScriptName, "Cannot find any organization sections in the settings file: " & $iniFile)
	Exit
EndIf

_ArraySort($allSections)
_ArrayColInsert($allSections, 1)
#EndRegion


#Region ==========================	 GUI	 ==========================
Local $mainGui = GUICreate("Infomonitors Screenshots Viewer", 300, 400)
GUISetFont(10)

Local $periodLabel = GUICtrlCreateLabel("Select organization(s) to view screenshots", 10, 12, 280, 20, $SS_CENTER)

Local $listView = GUICtrlCreateListView("Name", 10, 40, 280, 300, BitOr($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOCOLUMNHEADER))
GUICtrlSetState(-1, $GUI_FOCUS)
_GUICtrlListView_SetColumnWidth($listView, 0, 272)
_GUICtrlListView_AddArray($listView, $allSections)

Local $helpLabel = GUICtrlCreateLabel("Press and hold the ctrl or the shift key to select several lines", _
	10, 340, 280, 15, $SS_CENTER)
GUICtrlSetFont(-1, 8)
GUICtrlSetColor(-1, $COLOR_GRAY)

Local $selectAllButton = GUICtrlCreateButton("Select all", 10, 360, 120, 30)
Local $viewButton = GUICtrlCreateButton("View", 170, 360, 120, 30)
GUICtrlSetState(-1, $GUI_DISABLE)

GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")

Local $viewerGui = 0
Local $cGUI = 0
Local $needToExit = False
Local $childrenButtons[0][3]
Local $dW = 0
Local $dH = 0
Local $titleLabel = 0
Local $prevHour = 0
Local $nexHour = 0
Local $currentDate = 0
Local $currentHour = 0
Local $selectedItems[0][4]
Local $totalComputers = 0

While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $selectAllButton
			SelectAll()
		Case $viewButton
			$viewerGui = 0
			$cGUI = 0
			$needToExit = False
			$titleLabel = 0
			$prevHour = 0
			$nexHour = 0
			$currentDate = 0
			$currentHour = 0
			Local $tempArray[0][3]
			$childrenButtons = $tempArray
			Local $tempArray2[0][4]
			$selectedItems = $tempArray2
			$totalComputers = 0
			$dW = 0
			$dH = 0

			CreateChild()
	EndSwitch
	Sleep(20)
WEnd
#EndRegion


#Region ==========================	 Functions	 ==========================
Func MY_WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
	If  $wParam <> $listView Then Return

	Local $tagNMHDR, $event, $hwndFrom, $code
	$tagNMHDR = DllStructCreate("int;int;int", $lParam)
	If @error Then Return
	$event = DllStructGetData($tagNMHDR, 3)

	If  $event = $NM_CLICK Or $event = -12 Then
		CheckSelected()
	EndIf

	$tagNMHDR = 0
	$event = 0
	$lParam = 0
EndFunc


Func CreateChild()
	ToLog("CreateChild")
	_AD_Open()
	If @error Then
		MsgBox(16, "Infomonitors Screenshots Viewer", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)
		Return
	EndIf

	Local $sFQDN = _AD_SamAccountNameToFQDN()
	Local $iPos = StringInStr($sFQDN, ",")
	Local $sOU = StringMid($sFQDN, $iPos + 1)
	Local $aObjects[1][1]

	$currentHour = @HOUR
	Local $tempArray[0][4]
	$selectedItems = $tempArray
	$totalComputers = 0

	For $i = 0 To _GUICtrlListView_GetItemCount($listView)
		If Not _GUICtrlListView_GetItemSelected($listView, $i) Then ContinueLoop
		Local $currentOu[1][4]
		$currentOu[0][0] = _GUICtrlListView_GetItemTextString($listView, $i)
		$currentOu[0][1] = IniRead($iniFile, $currentOu[0][0], "screenshot_optional_path", "")
		$currentOu[0][2] = IniRead($iniFile, $currentOu[0][0], "main_ou", "")

		If $currentOu[0][2] Then
			Local $result = _AD_GetObjectsInOU($currentOu[0][2], "(&(objectCategory=computer)(name=*))", 2, "objectCategory,cn")

			If @error Then
				MsgBox(64, "Infomonitors Screenshots Viewer", "No OUs could be found for " & $currentOu[0][2])
			Else
				$currentOu[0][3] = $result
				$totalComputers += $result[0][0]
			EndIf
		EndIf

		_ArrayAdd($selectedItems, $currentOu)
	Next

	_AD_Close()

	$dW = @DesktopWidth
	$dH = @DesktopHeight - 40

	$hMonitor = GetMonitorFromPoint(0, 0)
	If $hMonitor <> 0 Then
		Local $arMonitorInfos[4]
		GetMonitorInfos($hMonitor, $arMonitorInfos)
		$dW = StringSplit($arMonitorInfos[1], ";",  $STR_NOCOUNT)[2]
		$dH = StringSplit($arMonitorInfos[1], ";",  $STR_NOCOUNT)[3]
	EndIf

	GUISetState(@SW_HIDE)

	Opt("GUIOnEventMode", 1)
	$viewerGui = GUICreate("Infomonitors Screenshots Results", 0, $dH, 0, 0)
	GUISetOnEvent($GUI_EVENT_CLOSE, "CloseChildrenWindow")
	GUISwitch($viewerGui)
	GUISetFont(10)

	WinMove("Infomonitors Screenshots Results", "", 0, 0, $dW, $dH)
;~ 	WinMove("Infomonitors Screenshots Results", "", 0, 0, 800, 600);$dW, $dH)

	$dW = WinGetClientSize("Infomonitors Screenshots Results")[0]
	$dH = WinGetClientSize("Infomonitors Screenshots Results")[1]

	GenerateChildViewingArea()

	While Not $needToExit
		Sleep(200)
	WEnd
EndFunc


Func GenerateChildViewingArea()
	If $titleLabel Then GUICtrlDelete($titleLabel)
	If $prevHour Then GUICtrlDelete($prevHour)
	If $nexHour Then GUICtrlDelete($nexHour)
	If $cGUI Then GUIDelete($cGUI)

	If $currentHour < 10 Then $currentHour = "0" & $currentHour

	GUISwitch($viewerGui)
	Local $progress = GUICtrlCreateProgress(10, 10, $dW - 20, 15)
	Local $progressLabel = GUICtrlCreateLabel("", 10, 25, $dW - 20, 15, $SS_CENTER)
	GUISetState(@SW_SHOW)

	$cGUI = GUICreate("Infomonitors Screenshots Iternal Window", $dW - 20, $dH - 60, 10, 50, $WS_CHILD, -1, $viewerGui)
	GUISetFont(10)

	Local $totalY = DrawData($selectedItems, $dW - 20, $dH - 60, $totalComputers, $progress, $progressLabel)
	_GuiScrollbars_Generate($cGUI, -1, $totalY, -1, 1, False, 180, True)

	GUISetState(@SW_SHOW)
	GUISwitch($viewerGui)
	GUICtrlDelete($progress)
	GUICtrlDelete($progressLabel)

	$titleLabel = GUICtrlCreateLabel("The screenshots creation time - " & $currentHour & ":00", $dW /2 - 340 / 2, _
		10, 340, 30, BitOR($SS_CENTERIMAGE, $SS_CENTER))
	GUICtrlSetFont(-1, 14, $FW_BOLD)

	Local $pos = ControlGetPos($viewerGui, "", $titleLabel)

	$prevHour = GUICtrlCreateButton("<<", $pos[0] - 40, 10, 30, 30)
	If $currentHour = 0 Then GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlSetOnEvent(-1, "PrevHour")

	$nexHour = GUICtrlCreateButton(">>", $pos[0] + $pos[2] + 10, 10, 30, 30)
	If $currentHour = @HOUR Then GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlSetOnEvent(-1, "HextHour")

	If UBound($childrenButtons, $UBOUND_ROWS) Then ControlFocus($cGUI, "", $childrenButtons[0][0])
EndFunc


Func PrevHour()
	$currentHour -= 1
	If $currentHour = 0 Then GUICtrlSetState($prevHour, $GUI_DISABLE)

	ToLog("PrevHour: " & $currentHour)

	GenerateChildViewingArea()
EndFunc


Func HextHour()
	$currentHour += 1
	If $currentHour = 23 Then GUICtrlSetState($nexHour, $GUI_DISABLE)

	ToLog("HextHour: " & $currentHour)

	GenerateChildViewingArea()
EndFunc


Func CloseChildrenWindow()
	ToLog("CloseChildrenWindow")

	Opt("GUIOnEventMode", 0)
	GUIDelete($viewerGui)
	GUISwitch($mainGui)
	SelectAll(False)
	GUICtrlSetState($viewButton, $GUI_DISABLE)
	GUISetState(@SW_SHOW)
	$needToExit = True
EndFunc


Func CompClicked()
	ToLog("CompClicked: " & @GUI_CtrlId)
	$detailsGUI = 0
	Local $index = _ArraySearch($childrenButtons, @GUI_CtrlId)
	If @error Then Return
	ToLog($childrenButtons[$index][0] & " - " & $childrenButtons[$index][1] & " - " & $childrenButtons[$index][2])
	GUISetState(@SW_HIDE, $cGUI)
	GUICtrlSetState($prevHour, $GUI_HIDE)
	GUICtrlSetState($titleLabel, $GUI_HIDE)
	GUICtrlSetState($nexHour, $GUI_HIDE)

	Local $detailsGUI = GUICreate("Infomonitors Screenshots Details", $dW, $dH, 0, 0, $WS_CHILD, -1, $viewerGui)
	GUISetFont(10)

	Local $ouName = $childrenButtons[$index][1]
	Local $compName = $childrenButtons[$index][2]
	Local $titleY = 40
	Local $dist = 10

	Local $compNameLabel = GUICtrlCreateLabel($compName & " at " & $currentHour & ":00", 140, 10, $dW - 280, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
	GUICtrlSetFont(-1, 14, $FW_BOLD)

	Local $mainPath = IniRead($iniFile, "general", "screenshot_path", "")
	Local $optionalPath = IniRead($iniFile, $ouName, "screenshot_optional_path", "")
	Local $strInStr = StringInStr(@ComputerName, $ouName)
	Local $rootPath = $strInStr ? $optionalPath : $mainPath
	Local $resultPath = $ouName & "\" & $currentDate & "\" & $currentHour & "\" & "*" & ".jpg"
	Local $path = $rootPath & StringReplace($resultPath, "*", $compName)

	Local $closeButton = GUICtrlCreateButton("Close", 10, 10, 120, 30)
	Local $tryRadmin = GUICtrlCreateButton("Connect via radmin", $dW - 120 - 10, 10, 120, 30)
	GUICtrlSetState(-1, $GUI_DISABLE)

	GUISetState(@SW_SHOW)

	If Not FileExists($path) Then _
		$path = ($strInStr ? $mainPath : $optionalPath) & StringReplace($resultPath, "*", $compName)

	Local $controlID
	If Not FileExists($path) Then
		$controlID = GUICtrlCreateLabel("Not exist", 10, 60, $dW - 20, $dH - 70, BitOr($SS_CENTER, $SS_CENTERIMAGE))
		GUICtrlSetBkColor(-1, $COLOR_RED)
		GUICtrlSetColor(-1, $COLOR_WHITE)
		GUICtrlSetFont(-1, 26, $FW_BOLD)
	Else
		ToLog("dW: " & $dW & " dH: " & $dH)
		ToLog("opW: " & $dW - 20 & " opH: " & $dH - 70)
		Local $optimalSize = GetOptimalControlSize($dW - 20, $dH - 70, $path)
		ToLog(_ArrayToString($optimalSize))

		$controlID = GUICtrlCreatePic($path, 10 + ($dW - 20 - $optimalSize[0]) / 2, 60 + ($dH - 70 - $optimalSize[1]) / 2, $optimalSize[0], $optimalSize[1], $SS_CENTERIMAGE)
		GUICtrlSetImage(-1, $path)
	EndIf

	GUISwitch($viewerGui)
	Opt("GUIOnEventMode", 0)

	Local $radminPath = ""
	If FileExists("C:\Program Files\Radmin Viewer 3\Radmin.exe") Then $radminPath = "C:\Program Files\Radmin Viewer 3\Radmin.exe"
	If FileExists("C:\Program Files (x86)\Radmin Viewer 3\Radmin.exe") Then $radminPath = "C:\Program Files (x86)\Radmin Viewer 3\Radmin.exe"

	Local $timeCounter = 2001
	While 1
		Local $msg = GUIGetMsg()
		If $msg = $GUI_EVENT_CLOSE Or $msg = $closeButton Then
			ToLog("$detailsGUI close")
			GUIDelete($detailsGUI)
			GUICtrlSetState($prevHour, $GUI_SHOW)
			GUICtrlSetState($titleLabel, $GUI_SHOW)
			GUICtrlSetState($nexHour, $GUI_SHOW)
			GUISetState(@SW_SHOW, $cGUI)
			If UBound($childrenButtons, $UBOUND_ROWS) Then ControlFocus($cGUI, "", $childrenButtons[0][0])
			ExitLoop
		ElseIf $msg = $tryRadmin Then
			ToLog("$tryRadmin")
			ShellExecute($radminPath, "/connect:" & $compName)
		EndIf

		If $timeCounter > 2000 Then
			Local $ping = Ping($compName, 2000)
			If @error = 1 Then $ping = "host is offline"
			If @error = 2 Then $ping = "host is unreachable"
			If @error = 3 Then $ping = "bad destination"
			If @error = 4 Then $ping = "other errors"
			If Not @error Then
				$ping &= " ms"
				If $radminPath Then GUICtrlSetState($tryRadmin, $GUI_ENABLE)
			Else
				GUICtrlSetState($tryRadmin, $GUI_DISABLE)
			EndIf

			GUICtrlSetData($compNameLabel, $compName & " at " & $currentHour & ":00, ping result: " & $ping)
			$timeCounter = 0
		EndIf

		Sleep(20)
		$timeCounter += 20
	WEnd

	Opt("GUIOnEventMode", 1)
EndFunc


Func GetOptimalControlSize($width, $height, $path)
	_GDIPlus_Startup()
	Local $image = _GDIPlus_ImageLoadFromFile($path)
	Local $imageWidth = _GDIPlus_ImageGetWidth($image)
	Local $imageHeight = _GDIPlus_ImageGetHeight($image)
	_GDIPlus_ImageDispose($image)
	_GDIPlus_Shutdown()

	ToLog("$imageWidth: " & $imageWidth & " $imageHeight: " & $imageHeight)
	ToLog("$width: " & $width & " $height: " & $height)

	Local $toReturn[2]

	Local $scaleFactor = $imageHeight / $imageWidth
	Local $controlHeight = $height
	Local $controlWidth = $imageWidth / ($imageHeight / $controlHeight)

	If $controlWidth > $width Then
		$controlHeight = $controlHeight / ($controlWidth / ($width))
		$controlWidth = $width
	EndIf

	$toReturn[0] = $controlWidth
	$toReturn[1] = $controlHeight

	Return $toReturn
EndFunc


Func DrawData($data, $width, $height, $total, $progress, $progressLabel)
	ToLog("---Drawing data---")
	ToLog("width: " & $width & " height: " & $height & " total: " & $total)
	Local $mainPath = IniRead($iniFile, "general", "screenshot_path", "")

	ToLog("Screenshot path: " & $mainPath)
	ToLog(_ArrayToString($data))

	$currentDate = @YEAR & @MON & @MDAY
	Local $imageSizeX = 200
	Local $imageSizeY = 160
	Local $dist = 10

	Local $countX = Floor($width / ($imageSizeX + $dist))
	ToLog("Max images on line: " & $countX)

	Local $totalWidth = $countX * $imageSizeX + ($countX - 1) * $dist
	ToLog("Total images width: " & $totalWidth)

	Local $startX = ($width - $totalWidth) / 2
	Local $nameY = 20
	Local $titleY = 40
	Local $currentY = 0
	Local $currentX = $startX
	Local $compCounter = 0

	For $i = 0 To UBound($data, $UBOUND_ROWS) - 1
		GUICtrlCreateLabel($data[$i][0], 0, $currentY, $width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
		GUICtrlSetFont(-1, 20, $FW_BOLD)
		GUICtrlSetBkColor(-1, $COLOR_SILVER)
		$currentY += $titleY + $dist

		Local $currentData = $data[$i][3]

		If Not UBound($currentData) Then
			ToLog("===ERROR EMPTY ARRAY===")
			GUICtrlCreateLabel("Cannot find the data in the active directory for: " & $data[$i][0], 0, _
				$currentY, $width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
			GUICtrlSetFont(-1, 14, $FW_BOLD)
			GUICtrlSetColor(-1, $COLOR_WHITE)
			GUICtrlSetBkColor(-1, $COLOR_RED)
			$currentY += $titleY
		EndIf

		Local $strInStr = StringInStr(@ComputerName, $data[$i][0])
		Local $rootPath = $strInStr ? $data[$i][1] : $mainPath
		Local $resultPath = $data[$i][0] & "\" & $currentDate & "\" & $currentHour & "\" & "*" & ".jpg"
		For $x = 1 To UBound($currentData) - 1
			Local $path = $rootPath & StringReplace($resultPath, "*", "_preview_" & $currentData[$x][1])

			If Not FileExists($path) Then _
				$path = ($strInStr ? $mainPath : $data[$i][1]) & StringReplace($resultPath, "*", "_preview_" & $currentData[$x][1])

			If Not FileExists($path) Then _
				$path = $rootPath & StringReplace($resultPath, "*", $currentData[$x][1])

			If Not FileExists($path) Then _
				$path = ($strInStr ? $mainPath : $data[$i][1]) & StringReplace($resultPath, "*", $currentData[$x][1])

			Local $control[3]
			$control[0] = GUICtrlCreateLabel("", $currentX - 1, $currentY - 1, $imageSizeX + 2, $imageSizeY + $nameY + 2, $SS_GRAYRECT)
			GUICtrlSetOnEvent(-1, "CompClicked")
			If Not FileExists($path) Then
				GUICtrlCreateLabel("Not exist", $currentX, $currentY, $imageSizeX, $imageSizeY, BitOR($SS_CENTER, $SS_CENTERIMAGE))
				GUICtrlSetBkColor(-1, $COLOR_RED)
				GUICtrlSetColor(-1, $COLOR_WHITE)
				GUICtrlSetFont(-1, 14, $FW_BOLD)
			Else
				Local $optimalSize = GetOptimalControlSize($imageSizeX, $imageSizeY, $path)
				GUICtrlCreateLabel("", $currentX, $currentY, $imageSizeX, $imageSizeY, $SS_WHITERECT)
				GUICtrlCreatePic($path, $currentX + ($imageSizeX - $optimalSize[0]) / 2, $currentY + ($imageSizeY - $optimalSize[1]) / 2, _
					$optimalSize[0], $optimalSize[1])
				GUICtrlSetImage(-1, $path)
			EndIf


			$control[1] = $data[$i][0]
			$control[2] = $currentData[$x][1]
			_ArrayTranspose($control)

			_ArrayAdd($childrenButtons, $control)

;~ 			GUICtrlSetOnEvent(-1, "CompClicked")

			GUICtrlCreateLabel($currentData[$x][1], $currentX, $currentY + $imageSizeY, $imageSizeX, $nameY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
			GUICtrlSetBkColor(-1, $COLOR_WHITE)

			$currentX += $imageSizeX + $dist
			If $currentX + $imageSizeX > $width - $startX Or $x = UBound($currentData) - 1 Then
				$currentX = $startX
				$currentY += $imageSizeY + $nameY + $dist
			EndIf

			$compCounter += 1
			GUICtrlSetData($progress, ($compCounter / $total) * 100)
			GUICtrlSetData($progressLabel, "Completed: " & $compCounter & " / " & $total & " current: " & $currentData[$x][1])
		Next
	Next

	Return $currentY
EndFunc


Func GetMonitorFromPoint($x, $y)
	Local $MONITOR_DEFAULTTONULL     = 0x00000000
    $hMonitor = DllCall("user32.dll", "hwnd", "MonitorFromPoint", _
                                            "int", $x, _
                                            "int", $y, _
                                            "int", $MONITOR_DEFAULTTONULL)
    Return $hMonitor[0]
EndFunc


Func GetMonitorInfos($hMonitor, ByRef $arMonitorInfos)
	Local $CCHDEVICENAME = 32
    Local $stMONITORINFOEX = DllStructCreate("dword;int[4];int[4];dword;char[" & $CCHDEVICENAME & "]")
    DllStructSetData($stMONITORINFOEX, 1, DllStructGetSize($stMONITORINFOEX))

    $nResult = DllCall("user32.dll", "int", "GetMonitorInfo", _
                                            "hwnd", $hMonitor, _
                                            "ptr", DllStructGetPtr($stMONITORINFOEX))
    If $nResult[0] = 1 Then
        $arMonitorInfos[0] = DllStructGetData($stMONITORINFOEX, 2, 1) & ";" & _
            DllStructGetData($stMONITORINFOEX, 2, 2) & ";" & _
            DllStructGetData($stMONITORINFOEX, 2, 3) & ";" & _
            DllStructGetData($stMONITORINFOEX, 2, 4)
        $arMonitorInfos[1] = DllStructGetData($stMONITORINFOEX, 3, 1) & ";" & _
            DllStructGetData($stMONITORINFOEX, 3, 2) & ";" & _
            DllStructGetData($stMONITORINFOEX, 3, 3) & ";" & _
            DllStructGetData($stMONITORINFOEX, 3, 4)
        $arMonitorInfos[2] = DllStructGetData($stMONITORINFOEX, 4)
        $arMonitorInfos[3] = DllStructGetData($stMONITORINFOEX, 5)
    EndIf

    Return $nResult[0]
EndFunc


Func SelectAll($set = True)
	_GUICtrlListView_BeginUpdate($listView)

	For $i = 0 To _GUICtrlListView_GetItemCount($listView)
		_GUICtrlListView_SetItemSelected($listView, $i, $set, $set)
	Next

	GUICtrlSetState($listView, $GUI_FOCUS)
	_GUICtrlListView_EndUpdate($listView)
	GUICtrlSetState($viewButton, $GUI_ENABLE)
EndFunc


Func CheckSelected()
	Local $ret = False
	For $i = 0 To _GUICtrlListView_GetItemCount($listView)
		If _GUICtrlListView_GetItemSelected($listView, $i) Then
		 $ret = True
		 GUICtrlSetState($viewButton, $GUI_ENABLE)
		 ExitLoop
		EndIf
	Next
	If Not $ret Then GUICtrlSetState($viewButton, $GUI_DISABLE)
EndFunc


Func ToLog($message)
	$message &= @CRLF
	ConsoleWrite($message)
EndFunc


Func HandleComError()
	ToLog(@ScriptName & " (" & $oMyError.scriptline & ") : ==> COM Error intercepted!" & @CRLF & _
		@TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oMyError.number) & @CRLF & _
		@TAB & "err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
		@TAB & "err.description is: " & @TAB & $oMyError.description & @CRLF & _
		@TAB & "err.source is: " & @TAB & @TAB & $oMyError.source & @CRLF & _
		@TAB & "err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
		@TAB & "err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF & _
		@TAB & "err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
		@TAB & "err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
		@TAB & "err.retcode is: " & @TAB & "0x" & Hex($oMyError.retcode) & @CRLF & @CRLF)
Endfunc
#EndRegion