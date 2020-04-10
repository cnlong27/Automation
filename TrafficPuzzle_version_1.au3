#include <Array.au3>
#include "HandleImgSearch.au3"
#include <ExcelCOM_UDF.au3>

;~ Sleep(4000)
;~ $a = PixelGetColor(634,290)
;~ $b = PixelGetColor(667,641)
;~ ConsoleWrite($a & @CRLF)
;~ ConsoleWrite($b & @CRLF)
;~ MsgBox(0,0,0)

$Handle = ""
;~ _GlobalImgInit($Handle, 0, 0, -1, -1, False, False, 15, 1000)

Global $g_bPaused = False
HotKeySet("{F1}", "TogglePause")
Func TogglePause()
		$g_bPaused = Not $g_bPaused
	While $g_bPaused
		Sleep(100)
		ToolTip('Script is "Paused"', 0, 0)
	WEnd
	ToolTip("")
 EndFunc   ;==>TogglePause

Local $posLevel1[6][2] = [[985, 620], [985, 620], [985, 620], [985, 620], [820, 395], [985, 620]]
Local $posLevel2[3][2] = [[990, 565], [990, 620], [990, 620]]
Local $posLevel3[4][2] = [[800, 615], [800, 615], [970, 620], [970, 560]]
Local $posLevel4[2][2] = [[800, 615], [1024, 360]]
Local $posLevel4_10Move[1][2] = [[985, 445]]
Local $unlimitedTime = 4000000

Local $hasData = False
Local $hasUnlimited = False

;800,615 ->Btn Play
;800,900 ->Tut Play lv2
;990,565 ->Right.3 8*8
;990,620 ->Right.4 8*8
;800,630 ->Next
;960,560 ->Right.3 7*8
;960,620 ->Right.4 7*8
;970,620 ->Right.4 8*7
;970,560 ->Right.3 8*7
;985,445 ->Right.1 8*8
;1024,360 ->Btn x(lv4)
;885,845 ->Btn lv1
;790,785 -> GetNow
;790,988 -> Btn Home


Local $simulatorList = WinList("NoxPlayer")

#Region small function

	Func _clickImage($image, $toren)
		$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\" & $image & ".bmp", 0, 0, -1, -1, $toren, 1000)
		If Not @error Then
			MouseMove($Result[1][0], $Result[1][1])
			Sleep(300)
			MouseClick('left', $Result[1][0], $Result[1][1])
		Else
			ConsoleWrite("_HandleImgSearch " & $image &": Fail" & @CRLF)
		EndIf
	EndFunc   ;==>_clickImage

	Func _clickImageWhile($image, $toren)
	$time = 1
	While $time = 1
		$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\" & $image & ".bmp", 0, 0, -1, -1, $toren, 1000)

		If $Result[0][0] == 0 Then
			$time = 1
		Else
			MouseMove($Result[1][0], $Result[1][1])
			Sleep(300)
			MouseClick('left', $Result[1][0], $Result[1][1])
			Sleep(1000)
			$time = 0
		EndIf
	WEnd
	EndFunc   ;==>_clickImageWhile

	Func _clickImageInTime($image, $toren, $time)
		For $i = 0 To $time
			_clickImage($image, $toren)
			Sleep(1000)
		Next
	EndFunc   ;==>_clickImageInTime

	Func MouseClickPoint($a, $b)
		MouseClick("left", $a, $b)
		Sleep(5000)
	EndFunc   ;==>MouseClickPoint
#EndRegion

#Region main function
	Func _installGame($email, $pass ,$mail)
		;esc login
		MouseClick('left', @DesktopWidth / 2, @DesktopHeight / 2)
		Sleep(10000)
		Send("{ESC}")
		Sleep(5000)

		; search game
		_clickImage('search01', 90)

		Sleep(3000)
		Send('traffic puzzle')
		Sleep(2000)
		Send("{ENTER}")
		Sleep(10000)

		;select game
	   $time = 1
	   While $time = 1
		 $Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\icongame.bmp", 0, 0, -1, -1, 120, 1000)

		 If $Result[0][0] == 0 Then
			MouseClick('left', 960, 190)
			Sleep(1000)
			Send("{ENTER}")
		 Else
			 MouseMove($Result[1][0], $Result[1][1])
			 Sleep(300)
			 MouseClick('left', $Result[1][0], $Result[1][1])
			 Sleep(1000)
			 $time = 0
		  EndIf
		  Sleep(7000)
	   WEnd



		Sleep(10000)
		;select CH play (1000,750 ,color = 16777215)
		If PixelGetColor(1000, 750) == 16777215 Then
			Sleep(500)
			Send("{DOWN}")
			Sleep(700)
			Send("{ENTER}")
			Sleep(1000)
			Send("{TAB}")
			Sleep(700)
			Send("{TAB}")
			Sleep(1000)
			Send("{ENTER}")
			Sleep(10000)
		EndIf

		;sign in CH play
		_clickImage('signin', 90)


		;sign in
		Sleep(15000)
		MouseClick('left', 940, 620)
		Sleep(1000)
		Send($email)
		Sleep(1000)
		Send("{ENTER}")
		Sleep(5000)

		;pass
		Sleep(2000)
		MouseClick('left', 940, 620)
		Sleep(1000)
		Send($pass)
		Sleep(1000)
		Send("{ENTER}")
		Sleep(5000)

		;verify code
		$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\verification code.bmp", 0, 0, -1, -1, 120, 1000)
		If $Result[0][0] <> 0 Then
			Sleep(1000)
			MouseMove(@DesktopWidth / 2, @DesktopHeight / 2)
			Sleep(500)
			MouseWheel('down',90)
			Sleep(2000)
			_clickImage("email", 120)
			Sleep(5000)
			MouseWheel('down',90)
			Sleep(2000)
			Send("{DOWN}")
			Sleep(1000)
			Send($mail)
			Sleep(1000)
			Send("{ENTER}")
			Sleep(5000)
		EndIf

		;add phone ???
		$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\addphone.bmp", 0, 0, -1, -1, 90, 1000)
		If $Result[0][0] <> 0 Then
			Sleep(1000)
			MouseClick('left', @DesktopWidth / 2, @DesktopHeight / 2)
			Sleep(1000)
			MouseWheel('down', 70)
			Sleep(1000)
			Send("{TAB}")
			Sleep(1000)
			Send("{ENTER}")
		EndIf

		;wellcome
		Sleep(5000)
		MouseClick("left", @DesktopWidth / 2 - 350, @DesktopHeight / 2 + 145)
		Sleep(700)
		Send("{TAB}")
		Sleep(700)
		Send("{TAB}")
		Sleep(700)
		Send("{TAB}")
		Sleep(700)
		Send("{TAB}")
		Sleep(1000)
		Send("{ENTER}")
		Sleep(10000)


		;google service
		MouseClick("left", @DesktopWidth / 2, @DesktopHeight / 2)
		MouseWheel('down', 15)
		Sleep(500)
		Send("{TAB}")
		Sleep(700)
		Send("{TAB}")
		Sleep(700)
		Send("{TAB}")
		Sleep(700)
		Send("{TAB}")
		Sleep(1000)
		Send("{ENTER}")
		Sleep(10000)

		;back
		_clickImage('back', 120)
		Sleep(7000)


		;search
		MouseClick('left', @DesktopWidth / 2, 190)
		Sleep(1500)
		Send('traffic puzzle')
		Sleep(1500)
		Send("{ENTER}")
		Sleep(15000)
		MouseMove(@DesktopWidth / 2, 250)

		$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\gameIcon.bmp", 0, 0, -1, -1, 90, 1000)
		If Not @error Then
			MouseMove($Result[1][0], $Result[1][1])
			Sleep(300)
			MouseClick('left', $Result[1][0], $Result[1][1])
		Else
		  MouseWheel('down', 80)
		  Sleep(2000)
		  ;into game
		  $temp = True
		  While $temp = True
			 $Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\intogame.bmp", 0, 0, -1, -1, 90, 1000)
			 If $Result[0][0] == 0 Then
				Sleep(3000)
				MouseWheel('down', 60)

			 Else
				MouseMove($Result[1][0], $Result[1][1])
				Sleep(300)
				MouseClick('left', $Result[1][0], $Result[1][1])
				Sleep(1000)
				$temp = False
			 EndIf
		  WEnd
		EndIf



		Sleep(10000)


		;install game
		_clickImage('install3', 90)
		Sleep(5000)


		;accept
		_clickImage('accept', 90)


		;open game
		_clickImageWhile('open', 90)

	EndFunc   ;==>_installGame
	Func _creatIdGame()
		;create Id game
		_clickImage('createId', 120)
		Sleep(10000)

		$color = PixelGetColor(@DesktopWidth / 2, @DesktopHeight / 2)
		While $color == 16777215
			;Gandt 01
			MouseClick('left', @DesktopWidth / 2, @DesktopHeight / 2)
			Sleep(700)
			Send("{TAB}")
			Sleep(700)
			Send("{TAB}")
			Sleep(700)
			Send("{TAB}")
			Sleep(1000)
			Send("{ENTER}")
			Sleep(4000)

			;Grant 02
			Send("{TAB}")
			Sleep(1000)
			Send("{ENTER}")
			Sleep(5000)

			;Allow (7 tab + enter)
			_clickImage('allow', 120)
			Sleep(10000)
			$color = PixelGetColor(@DesktopWidth / 2, @DesktopHeight / 2)
		WEnd

		GetRewards()

	EndFunc   ;==>_creatIdGame
	Func _createNox($times)
	   ShellExecute( @ScriptDir & "\Images\white.bmp")
	   Sleep(1000)
	   Send("{ENTER}")
	   Sleep(2000)
	   Run('C:\Program Files (x86)\Nox\bin\MultiPlayerManager.exe')
	   WinWait('Nox multi-instance manager','',10)
	   Sleep(5000)
	   WinMove('Nox multi-instance manager','',0,0)
	   Sleep(1000)
	    WinActivate('Nox multi-instance manager')
	   $temp = True
		For $i = 1 To $times
			MouseMove(765, 500)
			Sleep(1000)
			MouseClick('left', 765, 500)
			Sleep(1500)
			MouseClick('left', 570, 280)
			Sleep(2000)
			if $i = 1 Then
			   While $temp = True
				  $Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\playNox.bmp", 0, 0, -1, -1, 90, 1000)
				 If  $Result[0][0] <> 0 Then
					$temp = False
				 EndIf
			   WEnd
			EndIf

		Next
		MouseMove(200, 200) ;mouse move search image
		Sleep(5000)
	EndFunc   ;==>_createNox
	Func _OpenNox()

		WinActivate('Nox multi-instance manager')
		Sleep(2000)
		$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\playNox.bmp", 0, 0, -1, -1, 90, 1000)
	;~ 	_ArrayDisplay($Result)

		If Not @error Then
			For $i = 1 To $Result[0][0]
				MouseClick('left', $Result[$i][0], $Result[$i][1])
				Sleep(15000)
				WinActivate('Nox multi-instance manager')
				Sleep(2000)
			Next
		Else
			ConsoleWrite("Error Search openNox" & @CRLF)
		EndIf
	EndFunc   ;==>_OpenNox
	Func _closeNox()

	   Local $arrImg[] = ['quitNox', 'closeNox']

	   WinActivate('Nox multi-instance manager')
	   Sleep(2000)
   ;~ 	MsgBox(0,0,UBound($arrImg) - 1)
	   For $m = 0 To UBound($arrImg) - 1
   ;~ 		MsgBox(0,0,$arrImg[$m])
		   $Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\" & $arrImg[$m] & ".bmp", 0, 0, -1, -1, 90, 1000)
   ;~ 		_ArrayDisplay($Result)

		   If Not @error Then
			   For $i = 1 To $Result[0][0]
				   If $m = 1 Then
					   MouseClick('left', $Result[1][0] + 5, $Result[1][1] + 5)
				   Else
					   MouseClick('left', $Result[$i][0] + 5, $Result[$i][1] + 5)
				   EndIf
				   Sleep(2000)
				   Send("{ENTER}")
				   Sleep(10000)
				   WinActivate('Nox multi-instance manager')
				   Sleep(2000)
			   Next
		   Else
			   ConsoleWrite("Error Search " & $arrImg[$m] & @CRLF)
		   EndIf
	   Next
   ;~ MsgBox(0,0,0)
   EndFunc   ;==>_closeNox
	Func _closeNoxTimesOne()
	   Run('C:\Program Files (x86)\Nox\bin\MultiPlayerManager.exe')
	   WinWait('Nox multi-instance manager','',10)
	   Sleep(2000)
	   WinActivate('Nox multi-instance manager')
	   Sleep(2000)
	   $Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\closeNox.bmp", 0, 0, -1, -1, 90, 1000)
	   If Not @error Then
			 For $i = 1 To $Result[0][0]
				 MouseClick('left', $Result[1][0] + 5, $Result[1][1] + 5)
				 Sleep(2000)
				 Send("{ENTER}")
				 Sleep(4000)
				 WinActivate('Nox multi-instance manager')
				 Sleep(2000)
			 Next
		 Else
			 ConsoleWrite("Error Search closeNox.bmp" & @CRLF)
		 EndIf
	  EndFunc
    Func GetRewards()
	;congratulations
	$mau = 0
	For $i = 0 To 2
	  $mau = PixelGetColor(780, 250)
	  If $mau == 14632533 Then
		 MouseClick('left', 775, 685)
		 Sleep(7000)
		 MouseClick('left', 785, 860)
		 ExitLoop
	  EndIf
	  Sleep(5000)
	Next
;~ 	$mau = 0
;~ 	While $mau <> 14632533
;~ 		$mau = PixelGetColor(780, 250)
;~ 		If $mau == 14632533 Then
;~ 			MouseClick('left', 775, 685)
;~ 			Sleep(7000)
;~ 			MouseClick('left', 785, 860)
;~ 		EndIf
;~ 	WEnd
EndFunc

#EndRegion

#Region run function
	Func switchInstall()

		$oExcel = _ExcelBookOpen(@ScriptDir & '\gmail.xlsx')
		Sleep(1000)
		For $i = 1 To UBound($simulatorList) - 1
			WinActivate($simulatorList[$i][1]) ; Activate a window
			Sleep(3000)
			$email = _ExcelReadCell($oExcel, 'A' & $i)
	;~ 		MsgBox(0,0,$email)
			Sleep(500)
			$pass = _ExcelReadCell($oExcel, 'C' & $i)
			Sleep(500)
			$mail = _ExcelReadCell($oExcel, 'B' & $i)
	;~ 		MsgBox(0,0,$pass)
			_installGame($email, $pass,$mail)
			Sleep(2000) ; Wait 2 seconds
		Next

		If (UBound($simulatorList) - 1) <= 2 Then
			Sleep(25000)
		EndIf
	;~ 	_ExcelBookClose($oExcel)
	EndFunc   ;==>switchInstall

	Func switchPlayLevel1()
	  For $i = 1 To UBound($simulatorList) - 1
		 WinActivate($simulatorList[$i][1]) ; Activate a window
		 PlayLevel1First()
		 Sleep(2000) ; Wait 2 seconds
	  Next
	EndFunc   ;==>switchClickStart

	Func switchFindDataOrCreateID()
		For $i = 1 To UBound($simulatorList) - 1
			WinActivate($simulatorList[$i][1]) ; Activate a window
			Sleep(20000)
			$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\btnCont.bmp", 0, 0, -1, -1, 90, 1000)
			If $Result[0][0] <> 0 Then
			   $hasData = True
			   MouseMove($Result[1][0], $Result[1][1])
			   Sleep(500)
			   MouseClick('left', $Result[1][0], $Result[1][1])
			Else
			   $hasData = False
			   ConsoleWrite("_HandleImgSearch Continue Fail" & @CRLF)
			EndIf
			If $hasData == False Then
				_creatIdGame()
			EndIf
			Sleep(2000) ; Wait 2 seconds
		Next
	EndFunc   ;

	Func switchLoopLevel()
		For $a = 1 To 4
			For $i = 1 To UBound($simulatorList) - 1
				WinActivate($simulatorList[$i][1]) ; Activate a window
				playLevel1234()
				Sleep(2000) ; Wait 2 seconds
			Next
		Next
	EndFunc

   Func switchLoop60min()
	  For $i = 1 To UBound($simulatorList) - 1
		 CheckUnlimited()
		 If $hasUnlimited == True Then
			$begin = TimerInit()
			While 1
			   $dif = TimerDiff($begin)
			   If $dif > $unlimitedTime Then ExitLoop

			   WinActivate($simulatorList[$i][1]) ; Activate a window
			   playLevel1()
			   playLevel2()
			   playLevel3()
			   playLevel4()
			   Sleep(2000) ; Wait 2 seconds
			WEnd
		 EndIf
	  Next
   ;~
   EndFunc   ;==>switchLoop60min

   Func switchLoopLevel2to4()
	For $a = 1 To 5
		For $i = 1 To UBound($simulatorList) - 1
			WinActivate($simulatorList[$i][1]) ; Activate a window
			If $hasData == True Then
			   playLevel1()
			   playLevel2()
			   playLevel3()
			   playLevel4()
			ElseIf
			   playLevel2()
			   playLevel3()
			   playLevel4()
			   playLevel1()
			EndIf
			Sleep(2000) ; Wait 2 seconds
		Next
	Next
   EndFunc   ;==>switchLoopLevel2to4

	Func Automation()
	  _closeNoxTimesOne()
	  _createNox(2)
	  _OpenNox()
	  Sleep(10000)
	  $simulatorList = WinList("NoxPlayer")
	  Sleep(3000)
	  switchInstall()
	  switchPlayLevel1()
	  switchFindDataOrCreateID()
	  switchLoopLevel2to4()
	  switchLoop60min()
	  _closeNox()
	EndFunc   ;==>Automation

#EndRegion

#Region open and run Game
;~ _closeNoxTimesOne()
;~ _createNox(3)
;~ _OpenNox()
;~  switchInstall()
;~ _closeNox()
Automation()
#EndRegion

;open game from nox
;~ $Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\playgame.bmp", 0, 0, -1, -1, 120, 1000)
;~ _ArrayDisplay($Result)
;~ If not @error Then
;~ 	MouseMove($Result[1][0],$Result[1][1])
;~ 	Sleep(300)
;~ 	MouseClick('left',$Result[1][0],$Result[1][1])
;~ Else
;~ 	ConsoleWrite("_HandleImgSearch: Fail" & @CRLF)
;~ EndIf

#Region Play Level
Func PlayLevel1First()
	Sleep(10000)
	For $i = 0 To UBound($posLevel1) - 1
		MouseClickPoint($posLevel1[$i][0], $posLevel1[$i][1])
	Next
	Sleep(1000)
;~ search button NEXT
	_clickImageWhile("btnNext", 90)
EndFunc   ;==>PlayLevel1First

Func LoopLevel2to4()
	For $i = 0 To 4
		playLevel2()
		playLevel3()
		playLevel4()
		playLevel1()
	Next
EndFunc   ;==>LoopLevel2to4

Func playLevel1()
	_findPanelLuckyDay()
	_clickImageWhile("btnLv1", 90)
	Sleep(1000)
	_clickImageWhile("btnPlay", 90)
	Sleep(5000)
	For $i = 0 To UBound($posLevel1) - 1
		MouseClickPoint($posLevel1[$i][0], $posLevel1[$i][1])
	Next
	Sleep(1000)
;~ search button NEXT
	_clickImageWhile("btnNext", 90)
EndFunc   ;==>playLevel1

Func playLevel2()
	_findPanelLuckyDay()
	Sleep(1000)
	_clickImageWhile("btnPlay", 90)
	Sleep(2000)
	_clickImageWhile("btnPlayTut", 90)
	Sleep(2000)
	For $i = 0 To UBound($posLevel2) - 1
		MouseClickPoint($posLevel2[$i][0], $posLevel2[$i][1])
	Next
	_clickImageWhile("btnNext", 90)
EndFunc   ;==>playLevel2

Func playLevel3()
	_findPanelLuckyDay()
	Sleep(3000)
	_clickImageWhile("btnPlay", 90)
	Sleep(5000)
	For $i = 0 To UBound($posLevel3) - 1
		MouseClickPoint($posLevel3[$i][0], $posLevel3[$i][1])
	Next
	_clickImageWhile("btnNext", 90)
EndFunc   ;==>playLevel3

Func playLevel4()
	_findPanelLuckyDay()
	Sleep(1000)
	_clickImageWhile("btnPlay", 90)
	Sleep(3000)
	MouseClickPoint($posLevel4[0][0], $posLevel4[0][1])
	For $i = 0 To 10
		Sleep(1500)
		MouseClick("left", $posLevel4_10Move[0][0], $posLevel4_10Move[0][1])
	Next
	Sleep(1000)
	_clickImageWhile("btnClose", 90)
	Sleep(1000)
EndFunc   ;==>playLevel4

Func playLevel234()
	playLevel2()
	playLevel3()
	playLevel4()
 EndFunc   ;==>playLevel234

Func _findPanelLuckyDay()
	Sleep(5000)
	If PixelGetColor(634, 290) == 13313288 And PixelGetColor(667, 641) == 16184822 Then
		Sleep(2000)
		MouseClick("left", 790, 785)
		Sleep(3000)
		MouseClick("left", 790, 990)
		Sleep(3000)
	EndIf
EndFunc   ;==>_findPanelLuckyDay

Func CheckUnlimited()
	Sleep(3000)
	$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\unlimited.bmp", 0, 0, -1, -1, 90, 1000)
	If Not @error Then
		MouseMove($Result[1][0], $Result[1][1])
		Sleep(300)
		MouseClick('left', 790, 650)
		$hasUnlimited = True
	Else
		$hasUnlimited = False
	EndIf
 EndFunc   ;==>CheckUnlimited

#EndRegion

 Func _findImageBtn($image, $toren, $time)
	Sleep($time) ;Sleep 10s
	$Result = _HandleImgSearch($Handle, @ScriptDir & "\Images\" & $image & ".bmp", 0, 0, -1, -1, $toren, 1000)
	If Not @error Then
		$hasData = True
		MouseMove($Result[1][0], $Result[1][1])
		Sleep(300)
		MouseClick('left', $Result[1][0], $Result[1][1])
	Else
		$hasData = False
		;ConsoleWrite("_HandleImgSearch: Fail" & @CRLF)
	EndIf
EndFunc   ;==>_findImageBtn
