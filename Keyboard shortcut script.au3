#include <Date.au3>
#include <GUIConstantsEx.au3>

Global $email = "Saveļjevs, Mārtiņš (test.se)"
Global $possible = "Possible/True hit in DJ/WC" & @CRLF & "Name:" & @CRLF & "DOB:" & @CRLF & "Location:" & @CRLF & "Position:"
Global $Bogard = "test" & @CRLF & "Name:" & @CRLF & "DOB:" & @CRLF & "Role:" & @CRLF & "Role category:" & @CRLF & "Detailed category:" & @CRLF & "Country:" & @CRLF & "Start date:" & @CRLF & "End date:"
Global $FCRM = "test"
Global $Sorry = "test"

Global $isRunning = False
Global $areHotkeysEnabled = False

;;;;;;;;;; Create GUI
GUICreate("Email and Date Input Script", 200, 300)
Global $btnToggleScript = GUICtrlCreateButton("Start Script", 40, 20, 100, 40)
Global $lblShiftQ = GUICtrlCreateLabel("NumPad 3 - close the GUI", 40, 60, 150, 20)
Global $lblShiftX = GUICtrlCreateLabel("NumPad 2- stop the script", 40, 80, 150, 20)
Global $lblShiftE = GUICtrlCreateLabel("NumPad 4 - insert email", 40, 100, 150, 20)
Global $lblShiftG = GUICtrlCreateLabel("NumPad 5 - insert date", 40, 120, 150, 20)
Global $lblShiftA = GUICtrlCreateLabel("NumPad 8 - template about possibles", 40, 140, 150, 40)
Global $lblNumPad9 = GUICtrlCreateLabel("NumPad 7 - Template about test", 40, 170, 150, 80)
Global $lblShiftS = GUICtrlCreateLabel("NumPad 1 - Start Script", 40, 200, 150, 80)
Global $lblShiftS = GUICtrlCreateLabel("NumPad 9 - Template about not finding test", 40, 220, 150, 80)
Global $lblShift0 = GUICtrlCreateLabel("NumPad 0 - test", 40, 260, 150, 80)


GUISetState(@SW_SHOW)

HotKeySet("{NUMPAD2}", "ToggleScriptStop")  ;;;; ;; NumPad 2  to stop the script
HotKeySet("{NUMPAD4}", "InsertEmail")       ;;;; NumPad 4  to insert email
HotKeySet("{NUMPAD5}", "InsertDate")        ;;; NumPad 5  to insert today's date
HotKeySet("{NUMPAD3}", "EndScriptAndClose") ;;;;; NumPad 3  to end the script and close the GUI
HotKeySet("{NUMPAD8}", "InsertPossible")    ;;;;;; NumPad 8 template about possibles
HotKeySet("{NUMPAD7}", "Inserttest1") 		;;;;;;;; NumPad 7 triggers test
HotKeySet("{NUMPAD1}", "StartScript")       ;;;;;; NumPad 1  to start the script
HotKeySet("{NUMPAD9}", "InsertSorry") 		;;;;;; NumPad 9  to insert sorry comment
HotKeySet("{NUMPAD0}", "Inserttest2") 		;;;;;; NumPad 0  to insert test comment

While 1
    $msg = GUIGetMsg()
    Switch $msg
        Case $GUI_EVENT_CLOSE
            Exit
        Case $btnToggleScript
            ToggleScript()
    EndSwitch

    If $isRunning Then
        Sleep(100)
    EndIf
Wend

Func InsertEmail()
    If $isRunning And $areHotkeysEnabled Then
        Send($email)
    EndIf
EndFunc

Func InsertPossible()
    If $isRunning And $areHotkeysEnabled Then
        Send($possible)
    EndIf
EndFunc

Func InsertSorry()
    If $isRunning And $areHotkeysEnabled Then
        Send($Sorry)
    EndIf
EndFunc

Func Inserttest2()
    If $isRunning And $areHotkeysEnabled Then
        Send($Bogard)
    EndIf
EndFunc

Func Inserttest1()
    If $isRunning And $areHotkeysEnabled Then
        Send($FCRM)
    EndIf
EndFunc



Func InsertDate()
   If $isRunning And $areHotkeysEnabled Then
       Local $todaysDate = GetFormattedDate()
       Send($todaysDate)
    EndIf
EndFunc

Func GetFormattedDate()
    Local $todaysDate = _NowDate()
    Local $aDate = StringSplit($todaysDate, "/-")
    Return StringFormat("%04d-%02d-%02d", $aDate[3], $aDate[2], $aDate[1])
EndFunc



Func StartScript()
    If Not $isRunning Then
        ToggleScriptStart()
    EndIf
EndFunc

Func ToggleScriptStart()
    $isRunning = True
    GUICtrlSetData($btnToggleScript, "End Script")
    $areHotkeysEnabled = True
    While $isRunning
        Sleep(100)
    WEnd
EndFunc

Func ToggleScriptStop()
    $isRunning = False
    GUICtrlSetData($btnToggleScript, "Start Script")
    $areHotkeysEnabled = False
EndFunc

Func ToggleScript()
    If $isRunning Then
        ToggleScriptStop()
    Else
        ToggleScriptStart()
    EndIf
EndFunc

Func EndScriptAndClose()
    ToggleScriptStop()
    Exit
EndFunc

