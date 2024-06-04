;Excel data entry tool which scrapes values from a serial port dosimeter
;2024 Released into the Public Domain by JB
;Many parts pasted from forums, thank you
;com0com fake serial port software, and ComUDF were instrumental, thank you
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ComUDF.au3>
#include <Array.au3>
#include <GuiComboBox.au3>
#include <GuiListBox.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Debug.au3>
#include <WinAPIGdi.au3>
#include <APISysConstants.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
#include <WinAPI.au3>
#include <File.au3>
#include <Date.au3>

;$hWnd_Desktop = _WinAPI_GetDesktopWindow()
;$sKbLayout = _WinAPI_GetKeyboardLayout($hWnd_Desktop)
;;$hFUNC_KB = DllCallbackRegister('FUNC_KB', 'long', 'int;wparam;lparam') ;too sloooow
;;$hHook = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, DllCallbackGetPtr($hFUNC_KB), _WinAPI_GetModuleHandle(0))
Local $cellsearch='~*~*~*' ;look for all excel cells containing ***
Local $templatefilelist[1], $templatefolderlist[1]
Local $oWorkbook='', $hWndLast=0, $typedstring=''
Local $oExcel =_Excel_Open()
	If Not IsObj($oExcel) Then
		Exit MsgBox(0, "Error", "_Excel_Open()")
	EndIf
$aWorkBooks = _Excel_BookList()
$xlNormalView=1

Local $byrefBinary = 0					;response
Local $sComPorts = '', $sComPort='', $oldport='',$sDef=' baud=115200 data=8',$hComPort,$cell=0,$logging=0,$hLogFile,$sFolder,$sdosemultiplier=1000
Local $dose[1], $kV[1], $hvl[1],$serialnum,$time[1],$doserate[1], $pulse[1], $aCells[1]
Local $skV=" kV", $smR=" mR", $sms=" ms", $sHVL=" HVL", $sPulse=" Pulse"
Local $sinifile=@ScriptDir&'\'& StringRegExpReplace(@ScriptName,'.([^.]*$)','.ini'), $iResized = False, $xpos, $ypos, $winwidth,$winheight, $oldxpos, $reading=0,$numreading=0, $hWnd
Local $slogfile=@ScriptDir&'\'& StringRegExpReplace(@ScriptName,'.([^.]*$)','.log')
Local $reattach=0

Local $iScalex=1 ;=_WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth*1.33333333333333333333
Local $iScaley=1, $zoomscale ;=_WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth*1.355
;ConsoleWrite('Resolution Scaling X: '&$iScalex&' Y: '&$iScaley&@CRLF)

$logging = IniRead($sinifile, "Settings", "log",'')
if $logging=1 Then
	$hLogFile = FileOpen($slogfile, 1)
	_FileWriteLog($hLogFile,'Logging on: '&_NowCalcDate()&' '& _NowTime(4)&@CRLF)
ElseIf $logging='' Then
	IniWrite($sinifile, "Settings", "log", '1')
EndIf

$xpos = IniRead($sinifile, "Settings", "xpos","100")
$ypos = IniRead($sinifile, "Settings", "ypos","100")
$winwidth = IniRead($sinifile, "Settings", "winwidth","578")
$winheight = IniRead($sinifile, "Settings", "winheight","189")
$iScalex = Number(IniRead($sinifile, "Settings", "xscale","1"))
$iScaley= Number(IniRead($sinifile, "Settings", "yscale","1"))

$sComPort = IniRead($sinifile, "Settings", "port","1")
$oldxpos=$xpos

#Region ### START Koda GUI section ### Form=c:\koda\forms\unfors.kxf
$hWnd = GUICreate("Excel Data Entry for UNFORS", 578, 189, $xpos, $ypos, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_SIZEBOX));578, 189, 192, 124)
WinSetOnTop($hWnd, "", 1)
$Input1 = GUICtrlCreateInput("", 16, 8, 65, 28)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$Combo2 = GUICtrlCreateCombo("***", 96, 8, 65, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData($Combo2, "Yes|No|N/A|Y____|N____|N/A__|FIXED|P/T|MAN|'+3|'+2|'+1|'0|'-1|'-2|'-3|technique chart|console APR", "Item 2")
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$Combo3 = GUICtrlCreateCombo("", 256, 8, 224, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;browse combo
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$Combo4 = GUICtrlCreateCombo('', 496, 8, 65, 25, BitOR($CBS_DROPDOWN,$CBS_SIMPLE))
GUICtrlSetFont(-1, 11, 400, 0, "MS Sans Serif")
$Button1 = GUICtrlCreateButton(""&$skV, 16, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button2 = GUICtrlCreateButton(""&$smR, 96, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button3 = GUICtrlCreateButton(""&$sms, 176, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button4 = GUICtrlCreateButton(""&$sHVL, 256, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button5 = GUICtrlCreateButton(""&$sPulse, 336, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button6 = GUICtrlCreateButton("<", 416, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button7 = GUICtrlCreateButton(">", 496, 48, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button8 = GUICtrlCreateButton("| <", 16, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button9 = GUICtrlCreateButton("> |", 96, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button10 = GUICtrlCreateButton("First", 176, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button11 = GUICtrlCreateButton("Last", 256, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button12 = GUICtrlCreateButton("Rate", 336, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button13 = GUICtrlCreateButton("SN", 416, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button14 = GUICtrlCreateButton("Exit", 496, 88, 65, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Button15 = GUICtrlCreateButton("Yes", 176, 8, 65, 25)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1,$GUI_DOCKHCENTER+$GUI_DOCKVCENTER)
$Edit1 = GUICtrlCreateEdit("", 16, 128, 545, 49, BitOR($ES_AUTOVSCROLL,$ES_WANTRETURN,$WS_VSCROLL))
GUICtrlSetData($Edit1, '')
GUICtrlSetResizing(-1, $GUI_DOCKHCENTER+$GUI_DOCKBOTTOM)
GUICtrlSetPos($Edit1, 16, 128, 545, 49)
GUISetState(@SW_SHOWNOACTIVATE)
WinMove($hWnd,  '',$xpos, $ypos, $winwidth,$winheight)
updateports()
#EndRegion ### END Koda GUI section ###

if $iScalex=1 Then
	HotKeySet(']')
	HotKeySet('[')
	HotKeySet(';')
	HotKeySet("'")
	HotKeySet(']', 'HotKeyPressed')
	HotKeySet('[', 'HotKeyPressed')
	HotKeySet(';', 'HotKeyPressed')
	HotKeySet("'", 'HotKeyPressed')
EndIf


;GUIRegisterMsg($WM_SIZE, '_WM_SIZE')
;GUISetOnEvent($GUI_EVENT_RESIZED, '_WM_SIZE', $hWnd)
;GUIRegisterMsg($WM_ACTIVATE, "On_WM_ACTIVATE")
GUIRegisterMsg($WM_EXITSIZEMOVE, "on_WM_EXITSIZEMOVE")
$sFolder=IniRead($sinifile, "Settings", "Folder",'')
if $sFolder='' Then
	$sFolder=@UserProfileDir&'\Documents'
	IniWrite($sinifile, "Settings", "Folder", $sFolder)
EndIf
updatetemplates()

$sdosemultiplier=IniRead($sinifile, "Settings", "dosecorrectionfactor",'')
if $sdosemultiplier='' Then
	$sdosemultiplier=1000
	IniWrite($sinifile, "Settings", "dosecorrectionfactor", $sdosemultiplier)
EndIf

excelwindow('')

Do
   $buflen=_ComGetInputcount($hComPort)
	if $buflen>858 Then
		$d=_ComReadString($hComPort, $buflen, 1)
		parseserial($d)
		_ComClearOutputBuffer($hComPort)
		ControlFocus($hWnd, '', $Input1)
	Else
		_ComClearOutputBuffer($hComPort)
 	EndIf

	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_ComClosePort($hComPort)
			ExitLoop
			Case $GUI_EVENT_PRIMARYUP
			$fPos = WinGetPos($hWnd)
			if $fPos[1]<50 then
				GUISetState(@SW_HIDE,$hWnd)
			Else
				GUISetState(@SW_SHOWNOACTIVATE,$hWnd)
			EndIf
		Case $GUI_EVENT_RESIZED
			$fPos = WinGetPos($hWnd)
			if $fPos[1]<50 then
				GUISetState(@SW_HIDE,$hWnd)
			Else
			   GUISetState(@SW_SHOWNOACTIVATE,$hWnd)
			EndIf
		Case $Button14
			$nMsg = $GUI_EVENT_CLOSE
		Case $Button1
			updatecell(Number($kV[$reading]))
		Case $Button2
			updatecell(Number($dose[$reading]))
		Case $Button3
			updatecell(Number($time[$reading]))
		Case $Button4
			updatecell(Number($hvl[$reading]))
		Case $Button5
			updatecell(Number($pulse[$reading]))
		Case $Button12
			updatecell(Number($doserate[$reading]))
		Case $Button13
			updatecell($serialnum)
		Case $Button15
			updatecell('Yes')
			if $cell<UBound($aCells)-1 Then $cell+=1
			DrawCellOutline()
		Case $Button6
			if $reading>0 Then $reading-=1
			GUICtrlSetData($Edit1, ($reading)+1&': '&$kV[$reading]&', '&$dose[$reading]&', '&$time[$reading]&', '&$hvl[$reading]&', '&$pulse[$reading]&', '&$doserate[$reading]&', '&$serialnum)
			displayreading($reading)
			DrawCellOutline()
		Case $Button7
			if $reading<UBound($kV)-2 Then
				$reading+=1
				GUICtrlSetData($Edit1, ($reading)+1&': '&$kV[$reading]&', '&$dose[$reading]&', '&$time[$reading]&', '&$hvl[$reading]&', '&$pulse[$reading]&', '&$doserate[$reading]&', '&$serialnum)
				displayreading($reading)
			EndIf
			DrawCellOutline()
		Case $Button8
			$typedstring=GUICtrlRead($Input1)
			if $typedstring<>'' Then updatecell($typedstring)
			if $cell>0 Then
			   $typedstring=GUICtrlRead($Input1)
               if $typedstring<>'' Then updatecell($typedstring)
               $cell-=1
			   DrawCellOutline()
		    EndIf
		Case $Button9
			$typedstring=GUICtrlRead($Input1)
			if $typedstring<>'' Then updatecell($typedstring)
		    if $cell<UBound($aCells)-1 Then
				$cell+=1
				DrawCellOutline()
			EndIf
		Case $Button10
			$cell=0
			DrawCellOutline()
		Case $Button11
			$cell=UBound($aCells)-1
			DrawCellOutline()
		Case $Combo4
			$sComPort=StringReplace(GUICtrlRead($Combo4),'COM','')
			IniWrite($sinifile, "Settings", "port", $sComPort)
			updateports()
			excelwindow('') ;RELOAD CELL REFERENCES
		Case $Combo2
			updatecell(GUICtrlRead($Combo2))
		Case $Combo3
			pickfolder()
	EndSwitch

	if $reattach Then excelwindow('')
	Sleep(10)
Until $nMsg = $GUI_EVENT_CLOSE

;;_WinAPI_UnhookWindowsHookEx($hHook)
;;DllCallbackFree($hFUNC_KB)
OnAutoItExit()
Exit

Func _WM_SIZE($hWnd, $Msg, $wParam, $lParam)
	$iGUIWidth = BitAND($lParam, 0xFFFF)
	$iGUIHeight = BitShift($lParam, 16)
	$iResized = True
	ConsoleWrite("Function" & @CRLF)
	$fPos = WinGetPos($hWnd)
	$winwidth=$fPos[2]
	$winheight=$fPos[3]
		if $fPos[1]<50 then
			GUISetState(@SW_HIDE,$Edit1)
		Else
			GUISetState(@SW_SHOWNOACTIVATE,$Edit1)
		EndIf
    Return $GUI_RUNDEFMSG
EndFunc

func OnAutoItExit()
    if $logging Then FileClose($hLogFile)
    if  $xpos=$oldxpos Then Return
    IniWrite($sinifile, "Settings", "xpos", $xpos)
    IniWrite($sinifile, "Settings", "ypos", $ypos)
	IniWrite($sinifile, "Settings", "winwidth",$winwidth)
    IniWrite($sinifile, "Settings", "winheight",$winheight)
    if $sComPort<>$oldport Then IniWrite($sinifile, "Settings", "port", $sComPort)
EndFunc


Func on_WM_EXITSIZEMOVE($_hWnd, $msg, $wParam, $lParam)
    If $_hWnd = $hWnd Then
        $a = WinGetPos($hWnd)
		$xpos=$a[0]
		$ypos=$a[1]
		$winwidth=$a[2]
		$winheight=$a[3]
       ; DllCall("user32.dll", "long", "SetWindowPos", "hwnd", $hWnd, "hwnd", $hWnd_BOTTOM, "int", $a[0], "int", $wp[1], _
       ;         "int", 0, "int", 0, "long", BitOR($SWP_NOSIZE, $SWP_NOACTIVATE));BitOR($SWP_NOOWNERZORDER,$SWP_NOACTIVATE))
      	if $a[3]<170 then
				GUICtrlSetState($Edit1,$GUI_HIDE)
			Else
			   GUICtrlSetState($Edit1,$GUI_SHOW)
		   EndIf
   EndIf
    ConsoleWrite($xpos&' '&$ypos&@CRLF)
EndFunc

Func On_WM_ACTIVATE($hWnd, $msg, $wParam, $lParam)
;    Local $iState = BitAND($wParam, 0x0000FFFF), $iMinimize = BitShift($wParam, 16)
;    If $iState And $hWnd = $hGui Then
;        $wp = WinGetPos($hGui)
;        DllCall("user32.dll", "long", "SetWindowPos", "hwnd", $hWnd, "hwnd", $hWnd_BOTTOM, "int", $wp[0], "int", $wp[1], _
;                "int", $wp[2], "int", $wp[3], "long", $SWP_NOACTIVATE);BitOR($SWP_NOOWNERZORDER,$SWP_NOACTIVATE))
;    EndIf
EndFunc

Func updateports()
	_ComClosePort($hComPort)
	$sComPorts = _ComListPorts()
    if @error Then
		GUICtrlSetData($Edit1, 'FAILED to list COM ports')
		Return
	Else
		GUICtrlSetData($Edit1, 'OK, Listing COM Ports')
	EndIf
	GuiCtrlSetData($Combo4,  '','')
	GuiCtrlSetData($Combo4,  _ArrayToString($sComPorts),'')
	_GUICtrlComboBox_SetCurSel($Combo4, _ArraySearch($sComPorts,'COM'&$sComPort))
	$sComPort=StringReplace(GUICtrlRead($Combo4),'COM','')
	$hComPort = _ComOpenPort('COM'&$sComPort & $sDef)
	if @error Then
		GUICtrlSetData($Edit1, 'FAILED to Open COM'&$sComPort)
	Else
		GUICtrlSetData($Edit1, 'OK, Opened COM'&$sComPort)
	EndIf
EndFunc

Func parseserial($sDataStream)
	if StringInStr($sDataStream,'DataStream') then
		;_ComClearOutputBuffer($hComPort)
		;Return
	EndIf
	$valfound=0
	;ConsoleWrite($d&@CRLF)
	$sDataStream=BinaryToString($sDataStream)
	$a=StringRegExp($sDataStream,'.*</Value>|XiSerialNumber........',3)
	if @error Then Return ;ConsoleWrite('Regex Error '&@error&@CRLF)
		if $serialnum=0 Then
    	    For $j=0 to UBound($a)-1 Step 1
    		    if StringInStr($a[$j],'XiSerialNumber') Then
		    	    $serialnum=Number(StringRegExpReplace($a[$j],'.*="','',0)&@CRLF)
		        	ConsoleWrite('serialnum='&$serialnum&@CRLF)
		        	ExitLoop
	        	EndIf
        	Next
	    EndIf
	$reading=$numreading
	    For $j=UBound($a)-1 To 0 Step -1
	    	if StringInStr($a[$j],'"Dose"') Then
		    	$dose[$reading]=Number(StringRegExpReplace($a[$j],'.*">|</.*','',1)&@CRLF)*$sdosemultiplier
		    	ConsoleWrite('mR='&$dose[$numreading]&@CRLF)
				$valfound=1
		    	ExitLoop
	    	EndIf
    	Next

    	For $j=UBound($a)-1 To 0 Step -1
	    	if StringInStr($a[$j],'"kV"') Then
		    	$kV[$reading]=Number(StringRegExpReplace($a[$j],'.*">|</Value>','',2)&@CRLF)
		    	ConsoleWrite('kvp='&$kV[$numreading]&@CRLF)
		    	ExitLoop
	    	EndIf
    	Next

    	For $j=UBound($a)-1 To 0 Step -1
		    if StringInStr($a[$j],'"ExposureTime"') Then
		    	$time[$reading]=Number(StringRegExpReplace($a[$j],'.*">|</.*','',2)&@CRLF)*1000
		    	ConsoleWrite('time='&$time[$numreading]&@CRLF)
		    	ExitLoop
	    	EndIf
    	Next

	    For $j=UBound($a)-1 To 0 Step -1
		    if StringInStr($a[$j],'"HVL"') Then
			    $hvl[$reading]=Number(StringRegExpReplace($a[$j],'.*">|</.*','',2)&@CRLF)
		    	ConsoleWrite('HVL='&$hvl[$numreading]&@CRLF)
		    	ExitLoop
	    	EndIf
    	Next

    For $j=UBound($a)-1 To 0 Step -1
	    if StringInStr($a[$j],'"Pulse"') Then
		    $pulse[$reading]=Number(StringRegExpReplace($a[$j],'.*">|</.*','',2)&@CRLF)
		    ConsoleWrite('Pulse='&$pulse[$numreading]&@CRLF)
		    ExitLoop
	    EndIf
    Next

	For $j=UBound($a)-1 To 0 Step -1
	    if StringInStr($a[$j],'"DoseRate"') Then
		    $doserate[$reading]=Number(StringRegExpReplace($a[$j],'.*">|</.*','',2)&@CRLF)
		    ConsoleWrite('DoseRate='&$doserate[$numreading]&@CRLF)
		    ExitLoop
	    EndIf
	 Next

   ;_ArrayDisplay($kV)
   displayreading($numreading)

	If $valfound Then
		GUICtrlSetData($Edit1, $numreading&': '&$kV[$reading]&', '&$dose[$reading]&', '&$time[$reading]&', '&$hvl[$reading]&', '&$pulse[$reading]&', '&$doserate[$reading]&', '&$serialnum)
		if $logging=1 Then _FileWriteLog($hLogFile, 'LOGGED: '&_NowCalcDate()&' '& _NowTime(4)&$numreading&','&$kV[$reading]&','&$dose[$reading]&','&$time[$reading]&','&$hvl[$reading]&','&$pulse[$reading]&','&$doserate[$reading]&','&$serialnum&@CRLF&$sDataStream&@CRLF)

		$reading=$numreading;+1;UBound($kV)-1
		$numreading+=1
		_ArrayAdd($kV,0)
		_ArrayAdd($dose,0)
		_ArrayAdd($time,0)
		_ArrayAdd($HVL,0)
		_ArrayAdd($pulse,0)
		_ArrayAdd($doserate,0)
	EndIf
EndFunc

Func displayreading($n)
   GUICtrlSetData($Button1, round($kV[$n],1)&$skV)
   GUICtrlSetData($Button2, round($dose[$n],1)&$smR)
   GUICtrlSetData($Button3, round($time[$n],1)&$sms)
   GUICtrlSetData($Button4, round($HVL[$n],1)&$sHVL)
   GUICtrlSetData($Button5, round($pulse[$n],0)&$sPulse)
   GUICtrlSetData($Button12, round($doserate[$n],1))
   GUICtrlSetData($Button13, round($serialnum,0))
EndFunc

Func excelwindow($file)
	WinActivate("[CLASS:XLMAIN]")
	WinWaitActive("[CLASS:XLMAIN]", "", 10)
	Local $Activewindow=WinGetTitle("[ACTIVE]")
	ControlFocus( $hWnd, '', $Input1)
	WinActivate($hWnd)
	$cell=0
	;$Activewindow=StringReplace($Activewindow,' - Excel','') ;new excel
	;$Activewindow=StringRegExpReplace($Activewindow,'.*Excel - ','') ;old excel
	;MsgBox('','',$Activewindow)
	sleep(10)

;    $aWorkBooks = _Excel_BookList()
;_ArrayDisplay($aWorkBooks)
;    For $i = 0 To UBound($aWorkBooks) - 1
;      If StringInStr(String($aWorkBooks[$i][1]), $Activewindow) Then
;        $oWorkBook = _Excel_BookAttach($aWorkBooks[$i][2]&'\'&$aWorkBooks[$i][1]);
;		    If @error or StringInStr($Activewindow,$oWorkbook.Name)<1 Then
;               GUICtrlSetData($Edit1, "Can't Attach to Excel file: "&$aWorkBooks[$i][2]&'\'&$aWorkBooks[$i][1]&@CRLF&'Try closing excel, then starting this software, opening the spreadsheet, then pick a COM port.')
;	           Return
 ;           EndIf
  ;    EndIf
 ;   Next
  ; $oWorkbook = _Excel_BookAttach($Activewindow, "Title",'Default')
	if $file='' Then
		$oWorkbook = _Excel_BookAttach($Activewindow, "Title")
	Else
		$Activewindow=$templatefilelist[$file]
		$oWorkbook = _Excel_BookOpen($oExcel, $templatefolderlist[$file] & $templatefilelist[$file])
	EndIf
	If @error Then
		$oWorkbook = _Excel_BookOpen($oExcel, $Activewindow)
		If @error or StringInStr($Activewindow,$oWorkbook.Name)<1 Then
			GUICtrlSetData($Edit1, "Can't attach to an open file in "&$Activewindow&@CRLF&'Please pick a template first, or,'&@CRLF&'try closing excel, then starting this software, opening the spreadsheet, then pick a COM port.')
			Return
		EndIf
	EndIf

	if $file<>'' Then $aCells=StringSplit(IniRead($sinifile, "Files", $templatefilelist[$file],''),',')
	;_ArrayDisplay($aCells)
	Local $cellmode=''
	if UBound($aCells)-1<=1 Or $file='' Then
		$aCells = _Excel_RangeFind($oWorkbook, $cellsearch,"A1:R999",$xlFormulas,$xlWhole,True)
		If @error Then Return
		_ArrayColDelete($aCells, 0)
		_ArrayColDelete($aCells, 0)
		_ArrayColDelete($aCells, 3)
		_ArrayColDelete($aCells, 2)
		_ArrayColDelete($aCells, 1)
		$aCells=_ArrayToString($aCells,'')
		$aCells=StringReplace($aCells,@CRLF,',')
		$aCells=StringReplace($aCells,'$','')
		 if $file<>'' Then
		   if Not IniRead($sinifile, "Files", $templatefilelist[$file],'') then IniWrite($sinifile, "Files", $templatefilelist[$file],$aCells)
		 EndIf
		$cellmode='Found'
	Else
		$aCells=_ArrayToString($aCells,',',1)
		$cellmode='Loaded'
	EndIf

	GUICtrlSetData($Edit1, $cellmode&' cells: '&$aCells)
	$aCells=StringSplit($aCells,',',$STR_NOCOUNT)
	$oExcel.ActiveWindow.View=$xlNormalView
	DrawCellOutline()
EndFunc

Func DrawCellOutline()
	Local $X1,$Y1,$X2,$Y2,$interactive
	if UBound($aCells)-1<1 then Return
	;$oExcel.SendKeys('{ESC}') ;ESC exit cell editmode; wont work
	;If $oExcel.ActiveWindow.Interactive=True then Return  ;in edit mode, wont work
	;$oExcel.Interactive=False ;wont work
	_Excel_SheetList($oWorkbook)
	If @error Then Return ;hack to avoid editmode
	$zoomscale=$oExcel.ActiveWindow.Zoom/100
	$oWorkBook.ActiveSheet.Range($aCells[$cell]).Select
	;$oExcel.ActiveSheet.Shapes.AddShape(65, 0, 165,76.5, 25.5)
	$X1=$oExcel.ActiveWindow.PointsToScreenPixelsX($oExcel.Selection.Left*$iScalex*$zoomscale)
	$Y1=$oExcel.ActiveWindow.PointsToScreenPixelsY($oExcel.ActiveWindow.Selection.Top*$iScaley*$zoomscale)
	$iWidth=($oExcel.Selection.Width*$iScalex*$zoomscale)
	$iHeight=($oExcel.ActiveWindow.Selection.Height*$iScaley*$zoomscale)

	DrawBox($X1,$Y1, $iWidth, $iHeight, 5, 0x0000FF, 2000)
	ControlFocus( $hWnd, '', $Input1)
EndFunc

Func DrawBox($iStart_x, $iStart_y, $iWidth, $iHeight, $iThick, $iColor, $iTime)
    Local $hDC, $hPen, $o_Orig
    $hDC = _WinAPI_GetWindowDC(0) ; DC of entire screen (desktop)
    $hPen = _WinAPI_CreatePen($PS_SOLID, $iThick, $iColor)
    $o_Orig = _WinAPI_SelectObject($hDC, $hPen)
    _WinAPI_DrawLine($hDC, $iStart_x, $iStart_y,  $iStart_x+$iWidth,  $iStart_y)
	_WinAPI_DrawLine($hDC, $iStart_x, $iStart_y,  $iStart_x,  $iStart_y+$iHeight)
	_WinAPI_DrawLine($hDC, $iStart_x+$iWidth, $iStart_y+$iHeight,  $iStart_x,  $iStart_y+$iHeight)
    _WinAPI_DrawLine($hDC, $iStart_x+$iWidth,$iStart_y,  $iStart_x+$iWidth, $iStart_y+$iHeight)
    _WinAPI_SelectObject($hDC, $o_Orig)
    _WinAPI_DeleteObject($hPen)
    _WinAPI_ReleaseDC(0, $hDC)
EndFunc

Func _GetAppliedDPI()
    Local $AppliedDPI = RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "AppliedDPI")
    return $AppliedDPI
EndFunc

Func HotKeyPressed()
    Switch @HotKeyPressed
        Case "]"
			$iScalex=$iScalex*1.01
			ConsoleWrite('] '&$iScalex&@CRLF)
        Case "["
			$iScalex=$iScalex*0.99
			ConsoleWrite('[ '&$iScaley&@CRLF)
        Case ";"
			$iScaley=$iScaley*0.995
			ConsoleWrite('; '&$iScaley&@CRLF)
        Case "'"
			$iScaley=$iScaley*1.005
			ConsoleWrite("' "&$iScaley&@CRLF)
	EndSwitch
	DrawCellOutline()
	GUICtrlSetData($Edit1, ''&@CRLF&'xscale='&$iScalex&@CRLF&'yscale='&$iScaley)
	IniWrite($sinifile, "Settings", "xscale",$iScalex)
	IniWrite($sinifile, "Settings", "yscale",$iScaley)
EndFunc

Func updatecellnum()
	Local $celladdress=StringReplace($oExcel.Activecell.Address,'$','')
	Local $cellloc=_ArraySearch($aCells,$celladdress)
	if $cellloc>-1 Then $cell=$cellloc
	;ConsoleWrite('Skipping to cell '&$aCells[$cell]&' at location '&$cellloc&@CRLF)
EndFunc

Func updatecell($cellvalue)
	;WinWaitActive("[CLASS:XLMAIN]", "", 2)
	sleep(50) ;prevent crash
	;MsgBox('','',$cellvalue)
	_Excel_SheetList($oWorkbook)
	If @error Then Return ;hack to avoid editmode
	;If Not IsObj($oExcel.Interactive) Then Return
	_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $cellvalue, $oExcel.Activecell.Address)
	updatecellnum()
	GUICtrlSetData($Input1, '')
	ControlFocus($hWnd, '', $Input1)
EndFunc

Func FUNC_KB($nCode, $wParam, $lParam)
	Local Static $LastKey = Null
	If $nCode < 0 Then Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
	Local $tKEYHOOKS = DllStructCreate('dword vkCode;dword scanCode;dword flags;dword time;ulong_ptr dwExtraInfo', $lParam)
	Switch $wParam
        Case $WM_KEYDOWN
            If $tKEYHOOKS.vkCode = 0x1B Then $fExit = True
            If $LastKey = Null Then
                $LastKey = $tKEYHOOKS.vkCode
            EndIf
        Case $WM_KEYUP
            If $LastKey = $tKEYHOOKS.vkCode Then
                $LastKey = Null
				$typedstring=$typedstring&Chr(_WinAPI_MapVirtualKey($tKEYHOOKS.vkCode, $MAPVK_VK_TO_CHAR))
				ConsoleWrite($typedstring & @CRLF)
				updatecell($typedstring)
            EndIf
    EndSwitch
    Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
EndFunc

Func _LoWord($x)
    Return BitAND($x, 0xFFFF)
EndFunc   ;==>_LoWord

; http://msdn.microsoft.com/en-us/library/ms646307(VS.85).aspx
Func _WinAPI_MapVirtualKeyEx($sHexKey, $sKbLayout)
    Local Const $MAPVK_VK_TO_VSC = 0
    Local Const $MAPVK_VSC_TO_VK = 1
    Local Const $MAPVK_VK_TO_CHAR = 2
    Local Const $MAPVK_VSC_TO_VK_EX = 3
    Local Const $MAPVK_VK_TO_VSC_EX = 4

    Local $Ret = DllCall('user32.dll', 'long', 'MapVirtualKeyExW', 'int', '0x' & $sHexKey, 'int', 2, 'int', '0x' & $sKbLayout)
    Return $Ret[0]
EndFunc   ;==>_WinAPI_MapVirtualKeyEx

Func pickfolder()
	$choice=_GUICtrlComboBox_GetCurSel ($Combo3)
	if $choice>Ubound($templatefilelist)-1 OR _GUICtrlComboBox_GetCount($Combo3)=1 Then
        Local Const $sMessage = "Select a folder"
        Local $sFileSelectFolder = FileSelectFolder($sMessage, $sFolder)
        $sFolder=IniRead($sinifile, "Settings", "Folder",'')
        if $sFileSelectFolder<>$sFolder AND $sFileSelectFolder<>'' Then
           $sFolder=$sFileSelectFolder
         IniWrite($sinifile, "Settings", "Folder", $sFolder)
        EndIf
        updatetemplates()
	Else
        excelwindow($choice)
	EndIf
EndFunc

Func Search($current,$ext)
   Local $search = FileFindFirstFile($current & "\*.*")
   While 1
       Dim $file = FileFindNextFile($search)
       If @error Or StringLen($file) < 1 Then ExitLoop
       If Not StringInStr(FileGetAttrib($current & "\" & $file), "D") Then
      ; If StringRight($current & "\" & $file,StringLen($ext)) = $ext then
	  If StringInStr($file,$ext) AND StringLeft($file,1)<>'~' then
           ;FileWrite ("filelist.txt" , $current & "\" & $file & @CRLF)
		   ConsoleWrite($current & $file &@CRLF)
		   GUICtrlSetData($Edit1, 'Adding filelist to Templates, currently on: '&$current & $file&@CRLF&"Don't add root drives or large folders, as it will be slow"&@CRLF)
		   _ArrayAdd($templatefolderlist,$current& "\",'',',')
		   _ArrayAdd($templatefilelist,$file,'',',')
       Endif
       EndIf
       If StringInStr(FileGetAttrib($current & "\" & $file), "D") Then
       Search($current & "\" & $file, $ext)
       EndIf
   WEnd
   FileClose($search)
EndFunc

Func updatetemplates()
   for $a=UBound($templatefilelist)-1 to 0 Step -1
      _ArrayDelete($templatefolderlist, $a)
	  _ArrayDelete($templatefilelist, $a)
   Next
   Search($sFolder,'.xl')
   GuiCtrlSetData($Combo3,  '','')
   GUICtrlSetData($Combo3, _ArrayToString($templatefilelist, '|')&'|Browse for Template Folder','')
   _GUICtrlComboBox_SetCurSel($Combo3, Ubound($templatefilelist))
   ControlFocus( $hWnd, '', $Input1)
EndFunc