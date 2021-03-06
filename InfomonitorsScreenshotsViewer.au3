#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#pragma compile(ProductVersion, 0.3)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для отображения скриншотов инфомониторов)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555 - nn-admin@nnkk.budzdorov.su)
#pragma compile(ProductName, InfoscreenScreenshotsViewer)
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

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
#include <File.au3>


#Region ==========================	 Variables	 ==========================
Local $isGui = True
If $CmdLine[0] > 0 Then
   If $CmdLine[1] = "silent" Then $isGui = False
EndIf

Local $current_pc_name = @ComputerName
Local $errStr = "===ERROR=== "
Local $messageToSend = ""

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
$index = _ArraySearch($allSections, "mail_viewer")
If Not @error Then _ArrayDelete($allSections, $index)

If Not UBound($allSections) Then
	MsgBox($MB_ICONERROR, @ScriptName, "Cannot find any organization sections in the settings file: " & $iniFile)
	Exit
EndIf

_ArraySort($allSections)
If $isGui Then _ArrayColInsert($allSections, 1)

Local $mailSection = "mail_viewer"
Local $server_backup = "172.16.6.6"
Local $login_backup = "infoscreen_screenshots_viewer@nnkk.budzdorov.su"
Local $password_backup = "paqafapy"
Local $to_backup = "nn-admin@bzklinika.ru"
Local $send_email_backup = "1"

Local $server = IniRead($iniFile, $mailSection, "server", $server_backup)
Local $login = IniRead($iniFile, $mailSection, "login", $login_backup)
Local $password = IniRead($iniFile, $mailSection, "password", $password_backup)
Local $to = IniRead($iniFile, $mailSection, "to", $to_backup)
Local $send_email = IniRead($iniFile, $mailSection, "send_email", $send_email_backup)

Local $mainGui = 0
Local $switchTextError = "Show problems"
Local $switchTextAll = "Show all"
Local $selectAllButton = 0
Local $showOnlyError = 0
Local $viewButton = 0
Local $listView = 0
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
#EndRegion


If Not $isGui Then SilentModeCheckScreenshots()
FormMainGui()


Func FormMainGui()
	$mainGui = GUICreate("Infomonitors Screenshots Viewer", 300, 400)
	GUISetFont(10)

	Local $periodLabel = GUICtrlCreateLabel("Select organization(s) to view screenshots", 10, 12, 280, 20, $SS_CENTER)

	$listView = GUICtrlCreateListView("Name", 10, 40, 280, 270, BitOr($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOCOLUMNHEADER))
	GUICtrlSetState(-1, $GUI_FOCUS)
	_GUICtrlListView_SetColumnWidth($listView, 0, 272)
	_GUICtrlListView_AddArray($listView, $allSections)

	Local $helpLabel = GUICtrlCreateLabel("Press and hold the ctrl or the shift key to select several lines", _
		10, 310, 280, 15, $SS_CENTER)
	GUICtrlSetFont(-1, 8)
	GUICtrlSetColor(-1, $COLOR_GRAY)

	$showOnlyError = GUICtrlCreateCheckbox("Display only computers with problems", 10, 330, 280, 20)
	GUICtrlSetState(-1, $GUI_CHECKED)

	$selectAllButton = GUICtrlCreateButton("Select all", 10, 360, 120, 30)
	$viewButton = GUICtrlCreateButton("View", 170, 360, 120, 30)
	GUICtrlSetState(-1, $GUI_DISABLE)

	GUISetState(@SW_SHOW)
	GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				Exit
			Case $selectAllButton
				SelectAllButtonClicked()
			Case $viewButton
				ViewButtonClicked()
		EndSwitch

		If $needToView Then ViewButtonClicked()

		Sleep(20)
	WEnd
EndFunc


Func FormChildGui()
	ToLog("---FormChildGui---")

	Local $selectedItemsInList[0]
	For $i = 0 To _GUICtrlListView_GetItemCount($listView)
		If _GUICtrlListView_GetItemSelected($listView, $i) Then _
			_ArrayAdd($selectedItemsInList, _GUICtrlListView_GetItemTextString($listView, $i))
	Next

	$selectedItems = GetAdResultsForSelectedItems($selectedItemsInList)
	If Not UBound($selectedItems, $UBOUND_ROWS) Then
		ToLog("FormChildGui $selectedItems contains no rows")
		Return
	EndIf

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
	GUISetOnEvent($GUI_EVENT_CLOSE, "ChildrGuiCloseButtonClicked")
	GUISetFont(10)

	WinMove("Infomonitors Screenshots Results", "", 0, 0, $dW, $dH)

	$dW = WinGetClientSize("Infomonitors Screenshots Results")[0]
	$dH = WinGetClientSize("Infomonitors Screenshots Results")[1]

	CreateChildViewArea()

	While Not $needToExit
		Sleep(20)
	WEnd
EndFunc


Func FormComputerDetails()
	ToLog("---FormComputerDetails---")
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
			If $msg = $GUI_EVENT_CLOSE Then ChildrGuiCloseButtonClicked()
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





Func SelectAllButtonClicked($set = True)
	_GUICtrlListView_BeginUpdate($listView)

	For $i = 0 To _GUICtrlListView_GetItemCount($listView)
		_GUICtrlListView_SetItemSelected($listView, $i, $set, $set)
	Next

	GUICtrlSetState($listView, $GUI_FOCUS)
	_GUICtrlListView_EndUpdate($listView)
	GUICtrlSetState($viewButton, $GUI_ENABLE)
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

	FormChildGui()

	Opt("GUIOnEventMode", 0)
EndFunc


Func ChildrGuiCloseButtonClicked()
	ToLog("---ChildrGuiCloseButtonClicked---")
	GUIDelete($viewerGui)
	GUISwitch($mainGui)
	SelectAllButtonClicked(False)
	GUICtrlSetState($viewButton, $GUI_DISABLE)
	GUISetState(@SW_SHOW)
	$needToExit = True
EndFunc




Func CreateChildViewArea()
	ToLog("---CreateChildViewArea---")
	If $titleGUI Then GUIDelete($titleGUI)
	If $childGUI Then GUIDelete($childGUI)

	If StringLen($currentHour) < 2 Then $currentHour = "0" & $currentHour

	GUISwitch($viewerGui)
	Local $progress = GUICtrlCreateProgress(10, 10, $dW - 20, 15)
	Local $progressLabel = GUICtrlCreateLabel("", 10, 25, $dW - 20, 15, $SS_CENTER)
	GUISetState(@SW_SHOW)

	$childGUI = GUICreate("Infomonitors Screenshots Iternal Window", $dW - 20, $dH - 60, 10, 50, $WS_CHILD, -1, $viewerGui)
	GUISetFont(10)

	Local $onlyError = (GUICtrlRead($showOnlyError) = $GUI_CHECKED ? True : False)
	Local $totalY = CreateScreenshotsThumbinals($selectedItems, $dW - 20, $dH - 60, $totalComputers, $progress, $progressLabel, $onlyError)
	_GuiScrollbars_Generate($childGUI, -1, $totalY, -1, 1, False, 10, True)

	GUISetState(@SW_SHOW)
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
	GUICtrlSetOnEvent(-1, "ChildrGuiCloseButtonClicked")

	Local $switchView = GUICtrlCreateButton((GUICtrlRead($showOnlyError) = $GUI_CHECKED ? $switchTextAll : $switchTextError), _
		$dW - 130, 10, 120, 30)
	GUICtrlSetOnEvent(-1, "SwitchView")

	If UBound($childrenButtons, $UBOUND_ROWS) Then ControlFocus($childGUI, "", $childrenButtons[0][0])

	GUISetState(@SW_SHOW)
EndFunc


Func CreateScreenshotsThumbinals($data, $width, $height, $total, $progress, $progressLabel, $onlyError)
	ToLog("---Drawing data---")
	Local $tempArray[0][4]
	$childrenButtons = $tempArray
	Local $mainPath = IniRead($iniFile, "general", "screenshot_path", "")
	Local $infoscreenColor = IniRead($iniFile, "general", "infoscreen_standard_color", "")
	Local $infoscreenTimeTableColor = IniRead($iniFile, "general", "infoscreen_timetable_color", "")
	Local $desktop_color = IniRead($iniFile, "general", "desktop_color", "")
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
	Local $stringToReturn = ""

	For $i = 0 To UBound($data, $UBOUND_ROWS) - 1
		If $isGui Then
			GUICtrlCreateLabel($data[$i][0], 0, $currentY, $width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
			GUICtrlSetFont(-1, 20, $FW_BOLD)
			GUICtrlSetBkColor(-1, $COLOR_SILVER)
			$currentY += $titleY + $dist
		Else
			If $i Then $stringToReturn &= "|"
			$stringToReturn &= "---" & $data[$i][0] & "---|"
		EndIf

		Local $currentData = $data[$i][3]
		Local $infoscreenOu = IniRead($iniFile, $data[$i][0], "infoscreen_ou_to_run", "")
		Local $infoscreenTimetableOu = IniRead($iniFile, $data[$i][0], "infoscreen_timetable_ou_to_run", "")

		$infoscreenOu = StringSplit($infoscreenOu, ";", $STR_NOCOUNT)
		$infoscreenTimetableOu = StringSplit($infoscreenTimetableOu, ";", $STR_NOCOUNT)

		If Not UBound($currentData) Then
			Local $tempString = "Cannot find informations in the active directory"

			If Not $isGui Then
				$stringToReturn &= $tempString & "|"
				ContinueLoop
			EndIf

			GUICtrlCreateLabel($tempString, 0, $currentY, $width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
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

			If $isGui Then
				GUICtrlSetData($progress, ($compCounter / $total) * 100)
				GUICtrlSetData($progressLabel, "Completed: " & $compCounter & " / " & $total & " current: " & $currentData[$x][1])
			EndIf

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
							Local $tempString = "All computers are seems to be ok"
							If $isGui Then
								GUICtrlCreateLabel($tempString, 0, $currentY, $width, $titleY, BitOR($SS_CENTERIMAGE, $SS_CENTER))
								GUICtrlSetFont(-1, 14, $FW_BOLD)
								GUICtrlSetBkColor(-1, 0x00FF00)
								$currentY += $titleY + $dist
							Else
								$stringToReturn &= $tempString & "|"
							EndIf
						EndIf
					EndIf

					ContinueLoop
				EndIf
			EndIf

			$computersWithProblems += 1

			If Not $isGui Then
				$stringToReturn &= $currentData[$x][1]

				If Not FileExists($path) Then
					$stringToReturn &= " - not exist|"
				Else
					$stringToReturn &= " - not complies to normal state|"
				EndIf

				ContinueLoop
			EndIf

			Local $control[4]
			$control[0] = GUICtrlCreateLabel("", $currentX - 1, $currentY - 1, $imageSizeX + 2, $imageSizeY + $nameY + 2, $SS_GRAYRECT)
			GUICtrlSetOnEvent(-1, "FormComputerDetails")

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

	If Not $isGui Then Return $stringToReturn

	Return $currentY
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
		Local $optimalSize = GetOptimalControlSize($dW - 20, $dH - 60, $path)

		$startX = $startX + ($controlWidth - $optimalSize[0]) / 2
		$startY = $startY + ($controlHeight - $optimalSize[1]) / 2
		$controlWidth = $optimalSize[0]
		$controlHeight = $optimalSize[1]

		$newId[0] = GUICtrlCreateLabel("", $startX - 1, $startY - 1, $controlWidth + 2, $controlHeight + 2, $SS_GRAYRECT)
		GUICtrlSetState(-1, $GUI_HIDE)
		$newId[1] = GUICtrlCreatePic($path, $startX, $startY, $controlWidth, $controlHeight, $SS_CENTERIMAGE)
		GUICtrlSetState($newId[0], $GUI_SHOW)
	EndIf

	Return $newId
EndFunc




Func GetAdResultsForSelectedItems($array)
	If Not IsArray($array) Or Not Ubound($array) Then Return

	_AD_Open()
	If @error Then
		MsgBox(16, "GetAdResultsForSelectedItems", "Function _AD_Open encountered a problem. @error = " & _
			@error & ", @extended = " & @extended)
		Return
	EndIf

	Local $sFQDN = _AD_SamAccountNameToFQDN()
	Local $iPos = StringInStr($sFQDN, ",")
	Local $sOU = StringMid($sFQDN, $iPos + 1)
	Local $aObjects[1][1]

	$currentHour = @HOUR
	Local $tempArray[0][4]
	$totalComputers = 0

	For $i = 0 To UBound($array) - 1
		Local $currentOu[1][4]
		$currentOu[0][0] = $array[$i]
		$currentOu[0][1] = IniRead($iniFile, $currentOu[0][0], "screenshot_optional_path", "")
		$currentOu[0][2] = IniRead($iniFile, $currentOu[0][0], "main_ou", "")

		If $currentOu[0][2] Then
			Local $result = _AD_GetObjectsInOU($currentOu[0][2], "(&(objectCategory=computer)(name=*))", 2, "objectCategory,cn,distinguishedName")

			If @error Then
				MsgBox(64, "GetAdResultsForSelectedItems", "No OUs could be found for " & $currentOu[0][2])
			Else
				$currentOu[0][3] = $result
				$totalComputers += $result[0][0]
			EndIf
		EndIf

		_ArrayAdd($tempArray, $currentOu)
	Next

	_AD_Close()

	Return $tempArray
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


Func PrevHour()
	ToLog("---PrevHour---")
	$currentHour -= 1
	If $currentHour = 0 Then GUICtrlSetState($prevHour, $GUI_DISABLE)
	CreateChildViewArea()
EndFunc


Func HextHour()
	ToLog("---HextHour----")
	$currentHour += 1
	If $currentHour = 23 Then GUICtrlSetState($nexHour, $GUI_DISABLE)
	CreateChildViewArea()
EndFunc


Func SwitchView()
	ToLog("---SwitchView---")
	GUICtrlSetState($showOnlyError, GUICtrlRead($showOnlyError) = $GUI_CHECKED ? $GUI_UNCHECKED : $GUI_CHECKED)
	CreateChildViewArea()
EndFunc


Func GetOptimalControlSize($width, $height, $path, $checkError = False, $okColor = "", $errorColor = "")
	_GDIPlus_Startup()
	Local $image = _GDIPlus_ImageLoadFromFile($path)
	Local $imageWidth = _GDIPlus_ImageGetWidth($image)
	Local $imageHeight = _GDIPlus_ImageGetHeight($image)
	_GDIPlus_ImageDispose($image)
	_GDIPlus_Shutdown()

	If $checkError Then
		Local $pixelsColor[8]

		Local $tmp = StringReplace($path, "_preview_", "")
		If FileExists($tmp) Then $path = $tmp

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


Func SilentModeCheckScreenshots()
	$selectedItems = GetAdResultsForSelectedItems($allSections)
	Local $message = CreateScreenshotsThumbinals($selectedItems, 0, 0, 0, 0, 0, True)
	$messageToSend = StringReplace($message, "|", @CRLF)
	SendEmail()
EndFunc


Func SendEmail()
   If Not $send_email Then Exit

   Local $from = "Infoscreen screenshots viewer"
   Local $title = "Infosystems daily report"

   ToLog(@CRLF & "--- Sending email")
   If _INetSmtpMailCom($server, $from, $login, $to, _
		 $title, $messageToSend, "", "", "", $login, $password) <> 0 Then

	  _INetSmtpMailCom($server_backup, $from, $login_backup, $to_backup, _
		 $title, $messageToSend, "", "", "", $login_backup, $password_backup)
   EndIf

   Exit
EndFunc


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, _
   $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", _
   $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)

   Local $objEmail = ObjCreate("CDO.Message")
   Local $i_Error = 0
   Local $i_Error_desciption = ""

   $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
   $objEmail.To = $s_ToAddress

   If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
   If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

   $objEmail.Subject = $s_Subject

   If StringInStr($as_Body,"<") and StringInStr($as_Body,">") Then
	  $objEmail.HTMLBody = $as_Body
   Else
	  $objEmail.Textbody = $as_Body & @CRLF
   EndIf

   If $s_AttachFiles <> "" Then
	  Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
	  For $x = 1 To $S_Files2Attach[0] - 1
		 $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
		 If FileExists($S_Files2Attach[$x]) Then
			$objEmail.AddAttachment ($S_Files2Attach[$x])
		 Else
			$i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
			SetError(1)
			return 0
		 EndIf
	  Next
   EndIf

   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

   If $s_Username <> "" Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
   EndIf

   If $Ssl Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
   EndIf

   $objEmail.Configuration.Fields.Update
   $objEmail.Send

   if @error then
	  SetError(2)
   EndIf

   Return @error
EndFunc


Func ToLog($message)
	$message &= @CRLF
	ConsoleWrite($message)
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