#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 0.3)
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

Local $listView = GUICtrlCreateListView("Name", 10, 40, 280, 270, BitOr($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOCOLUMNHEADER))
GUICtrlSetState(-1, $GUI_FOCUS)
_GUICtrlListView_SetColumnWidth($listView, 0, 272)
_GUICtrlListView_AddArray($listView, $allSections)

Local $helpLabel = GUICtrlCreateLabel("Press and hold the ctrl or the shift key to select several lines", _
	10, 310, 280, 15, $SS_CENTER)
GUICtrlSetFont(-1, 8)
GUICtrlSetColor(-1, $COLOR_GRAY)

Local $showOnlyError = GUICtrlCreateCheckbox("Display only computers with problems", 10, 330, 280, 20)
GUICtrlSetState(-1, $GUI_CHECKED)

Local $selectAllButton = GUICtrlCreateButton("Select all", 10, 360, 120, 30)
Local $viewButton = GUICtrlCreateButton("View", 170, 360, 120, 30)
GUICtrlSetState(-1, $GUI_DISABLE)

GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")

Local $viewerGui = 0
Local $childGUI = 0
Local $titleGUI = 0
Local $needToExit = False
Local $childrenButtons[0][4]
Local $dW = 0
Local $dH = 0
Local $prevHour = 0
Local $nexHour = 0
Local $currentDate = 0
Local $currentHour = 0
Local $selectedItems[0][4]
Local $totalComputers = 0
Local $needToView = 0

Local $switchTextError = "Show problems"
Local $switchTextAll = "Show all"

While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $selectAllButton
			SelectAll()
		Case $viewButton
			ViewButtonClicked()
	EndSwitch

	If $needToView Then ViewButtonClicked()

	Sleep(20)
WEnd
#EndRegion


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


Func ViewButtonClicked()
	ToLog("---ViewButtonClicked---")
	$viewerGui = 0
	$childGUI = 0
	$needToExit = False
;~ 	$titleLabel = 0
	$prevHour = 0
	$nexHour = 0
;~ 	$switchView = 0
;~ 	$closeAllButton = 0
	$currentDate = 0
	$currentHour = 0
	Local $tempArray[0][4]
	$selectedItems = $tempArray
	$totalComputers = 0
	$dW = 0
	$dH = 0
	$needToView = 0

	CreateChild()

	Opt("GUIOnEventMode", 0)
EndFunc


Func MY_WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
	If  $wParam <> $listView Then Return

	Local $tagNMHDR, $event, $hwndFrom, $code
	$tagNMHDR = DllStructCreate("int;int;int", $lParam)
	If @error Then Return
	$event = DllStructGetData($tagNMHDR, 3)


	If $event = $NM_CLICK Or $event = -12 Then
		CheckSelected()
	ElseIf $event = $NM_DBLCLK Then
		$needToView = 1
	EndIf

	$tagNMHDR = 0
	$lParam = 0
	$event = 0
EndFunc


Func CreateChild()
	ToLog("---CreateChild---")
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
			Local $result = _AD_GetObjectsInOU($currentOu[0][2], "(&(objectCategory=computer)(name=*))", 2, "objectCategory,cn,distinguishedName")

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
	GUISetFont(10)
;~ 	GUISwitch($viewerGui)

	WinMove("Infomonitors Screenshots Results", "", 0, 0, $dW, $dH)
;~ 	WinMove("Infomonitors Screenshots Results", "", 0, 0, 800, 600);$dW, $dH)

	$dW = WinGetClientSize("Infomonitors Screenshots Results")[0]
	$dH = WinGetClientSize("Infomonitors Screenshots Results")[1]

	GenerateChildViewingArea()

	While Not $needToExit
		Sleep(20)
	WEnd
EndFunc


Func GenerateChildViewingArea()
	ToLog("---GenerateChildViewingArea---")
;~ 	If $titleLabel Then GUICtrlDelete($titleLabel)
;~ 	If $prevHour Then GUICtrlDelete($prevHour)
;~ 	If $nexHour Then GUICtrlDelete($nexHour)
;~ 	If $closeAllButton Then GUICtrlDelete($closeAllButton)
	If $titleGUI Then GUIDelete($titleGUI)
	If $childGUI Then GUIDelete($childGUI)

	If StringLen($currentHour) < 2 Then $currentHour = "0" & $currentHour

	GUISwitch($viewerGui)
	Local $progress = GUICtrlCreateProgress(10, 10, $dW - 20, 15)
	Local $progressLabel = GUICtrlCreateLabel("", 10, 25, $dW - 20, 15, $SS_CENTER)
	GUISetState(@SW_SHOW)

	$childGUI = GUICreate("Infomonitors Screenshots Iternal Window", $dW - 20, $dH - 60, 10, 50, $WS_CHILD, -1, $viewerGui)
	GUISetFont(10)

	Local $totalY = DrawData($selectedItems, $dW - 20, $dH - 60, $totalComputers, $progress, $progressLabel)
	_GuiScrollbars_Generate($childGUI, -1, $totalY, -1, 1, False, 10, True)

	GUISetState(@SW_SHOW)
;~ 	GUISwitch($viewerGui)
	GUICtrlDelete($progress)
	GUICtrlDelete($progressLabel)

	$titleGUI = GUICreate("Infomonitors Screenshots Window Title", $dW, 60, 0, 0, $WS_CHILD, -1, $viewerGui)
	GUISetFont(10)

	Local $titleLabel = GUICtrlCreateLabel("The screenshots creation time - " & $currentHour & ":00", ($dW - 340) / 2, _
		10, 340, 30, BitOR($SS_CENTERIMAGE, $SS_CENTER))
	GUICtrlSetFont(-1, 14, $FW_BOLD)

	Local $pos = ControlGetPos($viewerGui, "", $titleLabel)

	$prevHour = GUICtrlCreateButton("<<", $pos[0] - 40, 10, 30, 30)
	If $currentHour = 0 Then GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlSetOnEvent(-1, "PrevHour")

	$nexHour = GUICtrlCreateButton(">>", $pos[0] + $pos[2] + 10, 10, 30, 30)
	If $currentHour = @HOUR Then GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlSetOnEvent(-1, "HextHour")

	Local $closeAllButton = GUICtrlCreateButton("Close", 10, 10, 120, 30)
	GUICtrlSetOnEvent(-1, "CloseChildrenWindow")

	Local $switchView = GUICtrlCreateButton((GUICtrlRead($showOnlyError) = $GUI_CHECKED ? $switchTextAll : $switchTextError), _
		$dW - 130, 10, 120, 30)
	GUICtrlSetOnEvent(-1, "SwitchView")

	If UBound($childrenButtons, $UBOUND_ROWS) Then ControlFocus($childGUI, "", $childrenButtons[0][0])

	GUISetState(@SW_SHOW)
EndFunc


Func DrawData($data, $width, $height, $total, $progress, $progressLabel)
	ToLog("---Drawing data---")
	Local $tempArray[0][4]
	$childrenButtons = $tempArray
	Local $mainPath = IniRead($iniFile, "general", "screenshot_path", "")
	Local $infoscreenColor = IniRead($iniFile, "general", "infoscreen_standard_color", "")
	Local $infoscreenTimeTableColor = IniRead($iniFile, "general", "infoscreen_timetable_color", "")
	Local $desktop_color = IniRead($iniFile, "general", "desktop_color", "")
	Local $onlyError = (GUICtrlRead($showOnlyError) = $GUI_CHECKED ? True : False)
	$currentDate = @YEAR & @MON & @MDAY
	Local $imageSizeX = 200
	Local $imageSizeY = 160
	Local $dist = 10
	Local $countX = Floor($width / ($imageSizeX + $dist))
	Local $totalWidth = $countX * $imageSizeX + ($countX - 1) * $dist
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
		Local $infoscreenOu = IniRead($iniFile, $data[$i][0], "infoscreen_ou_to_run", "")
		Local $infoscreenTimetableOu = IniRead($iniFile, $data[$i][0], "infoscreen_timetable_ou_to_run", "")

		$infoscreenOu = StringSplit($infoscreenOu, ";", $STR_NOCOUNT)
		$infoscreenTimetableOu = StringSplit($infoscreenTimetableOu, ";", $STR_NOCOUNT)

		If Not UBound($currentData) Then
;~ 			ToLog("===ERROR EMPTY ARRAY===")
			GUICtrlCreateLabel("Cannot find informations in the active directory", 0, _
				$currentY, $width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
			GUICtrlSetFont(-1, 14, $FW_BOLD)
			GUICtrlSetColor(-1, $COLOR_WHITE)
			GUICtrlSetBkColor(-1, $COLOR_RED)
			$currentY += $titleY + $dist
			ContinueLoop
		EndIf

		Local $strInStr = StringInStr(@ComputerName, $data[$i][0])
		Local $rootPath = $strInStr ? $data[$i][1] : $mainPath
		Local $resultPath = $data[$i][0] & "\" & $currentDate & "\" & $currentHour & "\" & "*" & ".jpg"

		Local $computersWithProblems = 0

		For $x = 1 To UBound($currentData) - 1
			$compCounter += 1
			GUICtrlSetData($progress, ($compCounter / $total) * 100)
			GUICtrlSetData($progressLabel, "Completed: " & $compCounter & " / " & $total & " current: " & $currentData[$x][1])

			Local $path = $rootPath & StringReplace($resultPath, "*", "_preview_" & $currentData[$x][1])

			If Not FileExists($path) Then _
				$path = ($strInStr ? $mainPath : $data[$i][1]) & StringReplace($resultPath, "*", "_preview_" & $currentData[$x][1])

			If Not FileExists($path) Then _
				$path = $rootPath & StringReplace($resultPath, "*", $currentData[$x][1])

			If Not FileExists($path) Then _
				$path = ($strInStr ? $mainPath : $data[$i][1]) & StringReplace($resultPath, "*", $currentData[$x][1])

			Local $optimalSize = 0
			If FileExists($path) Then
				Local $okColor = ""

				If $onlyError Then
					For $ouName In $infoscreenOu
						If StringInStr($currentData[$x][2], $ouName) Then $okColor = $infoscreenColor
					Next

					For $ouName In $infoscreenTimetableOu
						If StringInStr($currentData[$x][2], $ouName) Then $okColor = $infoscreenTimeTableColor
					Next
				EndIf

				$optimalSize = GetOptimalControlSize($imageSizeX, $imageSizeY, $path, $onlyError, $okColor, $desktop_color)

				If $optimalSize = -1 And $onlyError Then
					If $x = UBound($currentData) - 1 Then
						$currentX = $startX
						If $computersWithProblems Then
							$currentY += $imageSizeY + $nameY + $dist
						Else
							GUICtrlCreateLabel("All computers are seems to be ok", 0, $currentY, _
								$width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
							GUICtrlSetFont(-1, 14, $FW_BOLD)
							GUICtrlSetBkColor(-1, 0x00FF00)
							$currentY += $titleY + $dist
						EndIf
					EndIf

					ContinueLoop
				EndIf
			EndIf

			Local $control[4]
			$control[0] = GUICtrlCreateLabel("", $currentX - 1, $currentY - 1, $imageSizeX + 2, $imageSizeY + $nameY + 2, $SS_GRAYRECT)
			GUICtrlSetOnEvent(-1, "CompClicked")

			If Not FileExists($path) Then
				GUICtrlCreateLabel("Not exist", $currentX, $currentY, $imageSizeX, $imageSizeY, BitOR($SS_CENTER, $SS_CENTERIMAGE))
				GUICtrlSetBkColor(-1, $COLOR_RED)
				GUICtrlSetColor(-1, $COLOR_WHITE)
				GUICtrlSetFont(-1, 14, $FW_BOLD)
			Else
				GUICtrlCreateLabel("", $currentX, $currentY, $imageSizeX, $imageSizeY, $SS_WHITERECT)
				Local $picX = $currentX + ($imageSizeX - $optimalSize[0]) / 2
				Local $picY = $currentY + ($imageSizeY - $optimalSize[1]) / 2
				Local $picW = $optimalSize[0]
				Local $picH = $optimalSize[1]
				GUICtrlCreatePic($path, $picX, $picY, $picW, $picH)
				GUICtrlSetImage(-1, $path)
			EndIf

			$control[3] = GUICtrlCreateLabel($currentData[$x][1], $currentX, $currentY + $imageSizeY, _
				$imageSizeX, $nameY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
			GUICtrlSetBkColor(-1, $COLOR_WHITE)

			$computersWithProblems += 1

			$control[1] = $data[$i][0]
			$control[2] = $currentData[$x][1]
			_ArrayTranspose($control)
			_ArrayAdd($childrenButtons, $control)

			$currentX += $imageSizeX + $dist
			If $currentX + $imageSizeX > $width - $startX Or $x = UBound($currentData) - 1 Then
				$currentX = $startX
				$currentY += $imageSizeY + $nameY + $dist
			EndIf
		Next
	Next

	Return $currentY
EndFunc


Func PrevHour()
	ToLog("---PrevHour---")
	$currentHour -= 1
	If $currentHour = 0 Then GUICtrlSetState($prevHour, $GUI_DISABLE)
	GenerateChildViewingArea()
EndFunc


Func HextHour()
	ToLog("---HextHour----")
	$currentHour += 1
	If $currentHour = 23 Then GUICtrlSetState($nexHour, $GUI_DISABLE)
	GenerateChildViewingArea()
EndFunc


Func SwitchView()
	ToLog("---SwitchView---")
	GUICtrlSetState($showOnlyError, GUICtrlRead($showOnlyError) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
	GenerateChildViewingArea()
EndFunc


Func CloseChildrenWindow()
	ToLog("---CloseChildrenWindow---")
	GUIDelete($viewerGui)
	GUISwitch($mainGui)
	SelectAll(False)
	GUICtrlSetState($viewButton, $GUI_DISABLE)
	GUISetState(@SW_SHOW)
	$needToExit = True
EndFunc


Func CompClicked()
	ToLog("---CompClicked---")
	$detailsGUI = 0
	Local $index = _ArraySearch($childrenButtons, @GUI_CtrlId)
	If @error Then Return
	GUICtrlSetFont($childrenButtons[$index][3], Default, $FW_BOLD)
	GUISetState(@SW_HIDE, $titleGUI)
	GUISetState(@SW_HIDE, $childGUI)

	Local $detailsGUI = GUICreate("Infomonitors Screenshots Details", $dW, $dH, 0, 0, $WS_CHILD, -1, $viewerGui)
	GUISetFont(10)

	Local $ouName = $childrenButtons[$index][1]
	Local $compName = $childrenButtons[$index][2]
	Local $dist = 10

	Local $showingHour = $currentHour

	Local $compNameLabel = GUICtrlCreateLabel($compName & " at " & $showingHour & ":00", ($dW - 340) / 2, 15, 340, 20, BitOR($SS_CENTERIMAGE, $SS_CENTER))
	GUICtrlSetFont(-1, 14, $FW_BOLD)

	Local $pingLabel = GUICtrlCreateLabel("Trying ping...", ($dW - 340) / 2, 34, 340, 15, BitOR($SS_CENTERIMAGE, $SS_CENTER))
	GUICtrlSetFont(-1, 8)
	GUICtrlSetColor(-1, $COLOR_GRAY)

	Local $mainPath = IniRead($iniFile, "general", "screenshot_path", "")
	Local $optionalPath = IniRead($iniFile, $ouName, "screenshot_optional_path", "")
	Local $strInStr = StringInStr(@ComputerName, $ouName)
	Local $rootPath = $strInStr ? $optionalPath : $mainPath
	Local $resultPath = $ouName & "\" & $currentDate & "\" & $showingHour & "\" & "*" & ".jpg"
	Local $path = $rootPath & StringReplace($resultPath, "*", $compName)

	Local $closeButton = GUICtrlCreateButton("Close", 10, 10, 120, 30)
	GUICtrlSetState(-1, $GUI_DISABLE)

	Local $tryRadmin = GUICtrlCreateButton("Connect via radmin", $dW - 120 - 10, 10, 120, 30)
	GUICtrlSetState(-1, $GUI_DISABLE)

	Local $position = ControlGetPos($detailsGUI, "", $compNameLabel)

	Local $previousButton = GUICtrlCreateButton("<<", $position[0] - 40, 10, 30, 30)
	If $showingHour = 0 Then GUICtrlSetState(-1, $GUI_DISABLE)

	Local $nextButton = GUICtrlCreateButton(">>", $position[0] + $position[2] + 10, 10, 30, 30)
	If $showingHour = @HOUR Then GUICtrlSetState(-1, $GUI_DISABLE)

	GUISetState(@SW_SHOW)

	If Not FileExists($path) Then _
		$path = ($strInStr ? $mainPath : $optionalPath) & StringReplace($resultPath, "*", $compName)

	Local $controlID = CreateControlForDetailView($path)

	GUISwitch($viewerGui)
	Opt("GUIOnEventMode", 0)

	Local $radminPath = ""
	If FileExists("C:\Program Files\Radmin Viewer 3\Radmin.exe") Then $radminPath = "C:\Program Files\Radmin Viewer 3\Radmin.exe"
	If FileExists("C:\Program Files (x86)\Radmin Viewer 3\Radmin.exe") Then $radminPath = "C:\Program Files (x86)\Radmin Viewer 3\Radmin.exe"

	Local $timeCounter = 1001
	While Not $needToExit
		Local $msg = GUIGetMsg()
		If $msg = $GUI_EVENT_CLOSE Or $msg = $closeButton Then
			ToLog("---DetailsGUI close---")
			GUIDelete($detailsGUI)
			GUISetState(@SW_SHOW, $childGUI)
			GUISetState(@SW_SHOW, $titleGUI)

			If UBound($childrenButtons, $UBOUND_ROWS) Then ControlFocus($childGUI, "", $childrenButtons[0][0])
			If $msg = $GUI_EVENT_CLOSE Then CloseChildrenWindow()
			ExitLoop
		ElseIf $msg = $tryRadmin Then
			ToLog("---Radmin---")
			ShellExecute($radminPath, "/connect:" & $compName)
		ElseIf $msg = $nextButton Or $msg = $previousButton Then
			ToLog("---Next or Previous---")
			GUICtrlDelete($controlID[0])
			GUICtrlDelete($controlID[1])

			$showingHour += ($msg = $nextButton) ? 1 : -1
			If StringLen($showingHour) < 2 Then $showingHour = "0" & $showingHour
			If $showingHour = 0 Then
				GUICtrlSetState($previousButton, $GUI_DISABLE)
			Else
				GUICtrlSetState($previousButton, $GUI_ENABLE)
			EndIf

			If $showingHour >= @HOUR Then
				GUICtrlSetState($nextButton, $GUI_DISABLE)
			Else
				GUICtrlSetState($nextButton, $GUI_ENABLE)
			EndIf

			GUICtrlSetData($compNameLabel, $compName & " at " & $showingHour & ":00")

			$resultPath = $ouName & "\" & $currentDate & "\" & $showingHour & "\" & "*" & ".jpg"
			$path = $rootPath & StringReplace($resultPath, "*", $compName)

			If FileExists Then
				$rootPath = $strInStr ? $mainPath : $optionalPath
				$path = $rootPath & StringReplace($resultPath, "*", $compName)
			EndIf

			GUISwitch($detailsGUI)
			$controlID = CreateControlForDetailView($path)
			GUISwitch($viewerGui)
		EndIf

		If $timeCounter > 1000 Then
			Local $ping = Ping($compName, 500)
			If @error = 1 Then $ping = "host is offline"
			If @error = 2 Then $ping = "host is unreachable"
			If @error = 3 Then $ping = "bad destination"
			If @error = 4 Then $ping = "other errors"
			If Not @error Then
				$ping &= " ms"
				If $radminPath And GUICtrlGetState($tryRadmin) >= $GUI_DISABLE Then GUICtrlSetState($tryRadmin, $GUI_ENABLE)
			Else
				GUICtrlSetState($tryRadmin, $GUI_DISABLE)
			EndIf

			GUICtrlSetData($pingLabel, "Ping results: " & $ping)

			If GUICtrlGetState($closeButton) >= $GUI_DISABLE Then GUICtrlSetState($closeButton, $GUI_ENABLE)
			$timeCounter = 0
		EndIf

		Sleep(20)
		$timeCounter += 20
	WEnd

	Opt("GUIOnEventMode", 1)
EndFunc


Func CreateControlForDetailView($path)
	ToLog("---CreateControlForDetailView---")
	Local $newId[2]

	Local $startX = 10
	Local $startY = 50
	Local $controlWidth = $dW - 20
	Local $controlHeight = $dH - 60
	If Not FileExists($path) Then
		$newId[0] = GUICtrlCreateLabel("", $startX - 1, $startY - 1, $controlWidth + 2, $controlHeight + 2, $SS_GRAYRECT)
		$newId[1] = GUICtrlCreateLabel("Not exist", $startX, $startY, $controlWidth, $controlHeight, BitOr($SS_CENTER, $SS_CENTERIMAGE))
		GUICtrlSetBkColor(-1, $COLOR_RED)
		GUICtrlSetColor(-1, $COLOR_WHITE)
		GUICtrlSetFont(-1, 26, $FW_BOLD)
	Else
;~ 		ToLog("dW: " & $dW & " dH: " & $dH)
;~ 		ToLog("opW: " & $dW - 20 & " opH: " & $dH - 70)
		Local $optimalSize = GetOptimalControlSize($dW - 20, $dH - 60, $path)
;~ 		ToLog(_ArrayToString($optimalSize))

		$startX = $startX + ($controlWidth - $optimalSize[0]) / 2
		$startY = $startY + ($controlHeight - $optimalSize[1]) / 2
		$controlWidth = $optimalSize[0]
		$controlHeight = $optimalSize[1]

		$newId[0] = GUICtrlCreateLabel("", $startX - 1, $startY - 1, $controlWidth + 2, $controlHeight + 2, $SS_GRAYRECT)
		GUICtrlSetState(-1, $GUI_HIDE)
		$newId[1] = GUICtrlCreatePic($path, $startX, $startY, $controlWidth, $controlHeight, $SS_CENTERIMAGE)
		GUICtrlSetState($newId[0], $GUI_SHOW)
;~ 		GUICtrlSetImage(-1, $path)
	EndIf

	Return $newId
EndFunc


Func GetOptimalControlSize($width, $height, $path, $checkError = False, $okColor = "", $errorColor = "")
	_GDIPlus_Startup()
	Local $image = _GDIPlus_ImageLoadFromFile($path)
	Local $imageWidth = _GDIPlus_ImageGetWidth($image)
	Local $imageHeight = _GDIPlus_ImageGetHeight($image)
	_GDIPlus_ImageDispose($image)
	_GDIPlus_Shutdown()

	If $checkError Then
;~ 		ToLog("CheckError")
		Local $pixelsColor[8]

		Local $tmp = StringReplace($path, "_preview_", "")
		If FileExists($tmp) Then $path = $tmp


;~ 		ToLog("GetOptimalControlSize: " & $path & " " & $checkError & " " & $okColor & " " & $errorColor)

		_GDIPlus_Startup()
		Local $imageToCheck = _GDIPlus_ImageLoadFromFile($path)

		Local $imageWidthCenter = _GDIPlus_ImageGetWidth($imageToCheck) / 2
		Local $imageHeightCenter = _GDIPlus_ImageGetHeight($imageToCheck) / 2

		$pixelsColor[0] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter, $imageHeightCenter)
		$pixelsColor[1] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter-20, $imageHeightCenter)
		$pixelsColor[2] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter+20, $imageHeightCenter)
		$pixelsColor[3] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter+13, $imageHeightCenter+7)
		$pixelsColor[4] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter, $imageHeightCenter-20)
		$pixelsColor[5] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter, $imageHeightCenter+20)
		$pixelsColor[6] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter-17, $imageHeightCenter-17)
		$pixelsColor[7] = _GDIPlus_BitmapGetPixel($imageToCheck, $imageWidthCenter+17, $imageHeightCenter+17)

		_GDIPlus_ImageDispose($imageToCheck)
		_GDIPlus_Shutdown()

		If $okColor Then
			$okColor = StringSplit($okColor, ";", $STR_NOCOUNT)
;~ 			_ArrayDisplay($okColor)

			For $color In $okColor
				If StringLen($color) = 6 Then
					If IsColorsComplies($pixelsColor, $color) Then Return -1
				EndIf
			Next
		ElseIf $errorColor Then
			$errorColor = StringSplit($errorColor, ";", $STR_NOCOUNT)

			For $color In $errorColor
				If StringLen($color) = 6 Then
					If Not IsColorsComplies($pixelsColor, $color) Then Return -1
				EndIf
			Next
		EndIf
	EndIf

;~ 	ToLog("$imageWidth: " & $imageWidth & " $imageHeight: " & $imageHeight)
;~ 	ToLog("$width: " & $width & " $height: " & $height)

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


Func IsColorsComplies($rawColorsArray, $colorToCheck)
	For $rawColor In $rawColorsArray
		$rawColor = Hex($rawColor, 6)

		Local $curCol[3]
		$curCol[0] = StringLeft($rawColor, 2)
		$curCol[1] = StringMid($rawColor, 2, 2)
		$curCol[2] = StringRight($rawColor, 2)

		Local $checkCol[3]
		$checkCol[0] = StringLeft($colorToCheck, 2)
		$checkCol[1] = StringMid($colorToCheck, 2, 2)
		$checkCol[2] = StringRight($colorToCheck, 2)

;~ 		ToLog("cur: " & _ArrayToString($curCol) & " check: " & _ArrayToString($checkCol))

		Local $colDiff = 0
		For $x = 0 To 2
			$colDiff += Abs(Dec($curCol[$x]) - Dec($checkCol[$x]))
		Next

		If $colDiff < 60 Then Return True
	Next

	Return False
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