#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <Date.au3>

;;;;;;;;;;;;;;;;;;;;;;; Initialize global variables
Global $oOutlook, $oNamespace, $oCustomerScreeningFolder, $oMailbox
Global $aCategories[5], $iCounter = 0, $aCategorizedEmails, $hGUI
Global $bCategorizing = False, $bRunning = False
Global $sLogFile = @ScriptDir & "\Categorizer.log" ; Log file path

;;;;;;;;;;;;;;;;;;;; Function to create and display the GUI
Func CreateGUI()
    $hGUI = GUICreate("Categorizer", 400, 350)

    GUICtrlCreateLabel("Person 1 Category:", 10, 20)
    GUICtrlCreateLabel("Person 2 Category:", 10, 60)
    GUICtrlCreateLabel("Person 3 Category:", 10, 100)
    GUICtrlCreateLabel("Person 4 Category:", 10, 140)
    GUICtrlCreateLabel("Person 5 Category:", 10, 180)

    Global $hCat1 = GUICtrlCreateInput("", 130, 20, 200, 20)
    Global $hCat2 = GUICtrlCreateInput("", 130, 60, 200, 20)
    Global $hCat3 = GUICtrlCreateInput("", 130, 100, 200, 20)
    Global $hCat4 = GUICtrlCreateInput("", 130, 140, 200, 20)
    Global $hCat5 = GUICtrlCreateInput("", 130, 180, 200, 20)

    Global $hRemoveCat1 = GUICtrlCreateButton("Remove", 340, 20, 50, 20)
    Global $hRemoveCat2 = GUICtrlCreateButton("Remove", 340, 60, 50, 20)
    Global $hRemoveCat3 = GUICtrlCreateButton("Remove", 340, 100, 50, 20)
    Global $hRemoveCat4 = GUICtrlCreateButton("Remove", 340, 140, 50, 20)
    Global $hRemoveCat5 = GUICtrlCreateButton("Remove", 340, 180, 50, 20)

    Global $aCatInputs[5] = [$hCat1, $hCat2, $hCat3, $hCat4, $hCat5]
    Global $aRemoveButtons[5] = [$hRemoveCat1, $hRemoveCat2, $hRemoveCat3, $hRemoveCat4, $hRemoveCat5]

    Global $hOkButton = GUICtrlCreateButton("OK", 50, 240, 100, 30)
    Global $hPauseButton = GUICtrlCreateButton("Pause", 150, 240, 100, 30)
    Global $hEndButton = GUICtrlCreateButton("End", 250, 240, 100, 30)

    GUISetState(@SW_SHOW)

    ;;;;;;;;;;;;;;;;;;;; Main GUI message loop
    While 1
        Local $nMsg = GUIGetMsg()

        Select
            Case $nMsg = $GUI_EVENT_CLOSE
                StopCategorizing()
                GUIDelete($hGUI)
                Exit

            Case $nMsg = $hEndButton
                StopCategorizing()
                GUIDelete($hGUI)
                Exit

            Case $nMsg = $hPauseButton
                If $bCategorizing Then
                    PauseCategorizing()
                    GUICtrlSetData($hPauseButton, "Resume")
                Else
                    ResumeCategorizing()
                    GUICtrlSetData($hPauseButton, "Pause")
                EndIf

            Case $nMsg = $hOkButton
                MsgBox($MB_SYSTEMMODAL, "Script Status", "Script is starting")

                For $i = 0 To 4
                    $aCategories[$i] = GUICtrlRead($aCatInputs[$i])
                Next

                If Not $bRunning Then
                    $bRunning = True
                    StartCategorizing()
                ElseIf Not $bCategorizing Then
                    StartCategorizing()
                EndIf
        EndSelect
    WEnd
EndFunc

Func RemoveCategory($iIndex)
    UnmarkEmails($aCategories[$iIndex])

    $aCategories[$iIndex] = ""
    GUICtrlSetData($aCatInputs[$iIndex], "")
EndFunc

Func UnmarkEmails($sCategory)
    $oOutlook = ObjCreate("Outlook.Application")
    $oNamespace = $oOutlook.GetNamespace("MAPI")

    $oMailbox = $oNamespace.Folders("Customer Screening")
    If Not IsObj($oMailbox) Then
        MsgBox(0, "Error", "Unable to access the mailbox 'Customer Screening'. Exiting script.")
        CleanupCOM()
        Exit
    EndIf

    $oCustomerScreeningFolder = $oMailbox.Folders("Inbox")
    If Not IsObj($oCustomerScreeningFolder) Then
        MsgBox(0, "Error", "Unable to access the 'Inbox' folder in 'Customer Screening'. Exiting script.")
        CleanupCOM()
        Exit
    EndIf

    Local $oItems = $oCustomerScreeningFolder.Items

    For $iIndex = 1 To $oItems.Count
        Local $oMail = $oItems.Item($iIndex)

        If StringInStr($oMail.Categories, $sCategory) Then
            $oMail.Categories = ""
            $oMail.Save()

            If $aCategorizedEmails.Exists($oMail.EntryID) Then
                $aCategorizedEmails.Remove($oMail.EntryID)
            EndIf
        EndIf
        $oMail = Null
    Next
    CleanupCOM()
EndFunc

Func StartCategorizing()
    $bCategorizing = True
    If Not IsObj($aCategorizedEmails) Then
        $aCategorizedEmails = ObjCreate("Scripting.Dictionary")
    EndIf
    AdlibRegister("CategorizeEmails", 5000)

    WriteLog("Script is starting")
EndFunc

Func PauseCategorizing()
    $bCategorizing = False
    AdlibUnRegister("CategorizeEmails")
EndFunc

Func ResumeCategorizing()
    $bCategorizing = True
    AdlibRegister("CategorizeEmails", 5000)
EndFunc

Func StopCategorizing()
    $bCategorizing = False
    $bRunning = False
    AdlibUnRegister("CategorizeEmails")

    WriteLog("Script stopped")

    MsgBox($MB_SYSTEMMODAL, "Script Status", "Script stopped")
EndFunc

Func CategorizeEmails()
    $oOutlook = ObjCreate("Outlook.Application")
    $oNamespace = $oOutlook.GetNamespace("MAPI")

    $oMailbox = $oNamespace.Folders("test")
    If Not IsObj($oMailbox) Then
        MsgBox(0, "Error", "Unable to access the mailbox 'Customer Screening'. Exiting script.")
        CleanupCOM()
        Exit
    EndIf

    $oCustomerScreeningFolder = $oMailbox.Folders("test")
    If Not IsObj($oCustomerScreeningFolder) Then
        MsgBox(0, "Error", "Unable to access the 'Inbox' folder in 'Customer Screening'. Exiting script.")
        CleanupCOM()
        Exit
    EndIf

    Local $oItems = $oCustomerScreeningFolder.Items
    $oItems.Sort("[ReceivedTime]", False)

    Local $bCategorized = False
    Local $iProcessedCount = 0
    Local $iBatchSize = 50

    For $iIndex = 1 To $oItems.Count
        Local $oMail = $oItems.Item($iIndex)
        Local $sEntryID = $oMail.EntryID

        ;;;;;;;;;; Checks if the email is unread, has a PDF attachment, has not been replied or forwarded,
        If $oMail.UnRead And $oMail.Categories = "" And HasPDFAttachment($oMail) And Not IsRepliedOrForwarded($oMail) And Not $aCategorizedEmails.Exists($sEntryID) Then
            While $aCategories[$iCounter] = ""
                $iCounter = Mod($iCounter + 1, 5)
            Wend

            $oMail.Categories = $aCategories[$iCounter]
            $oMail.Save()

            $aCategorizedEmails.Add($sEntryID, 1)

            $iCounter = Mod($iCounter + 1, 5)
            $bCategorized = True
            $iProcessedCount += 1

            If $iProcessedCount >= $iBatchSize Then
                ExitLoop
            EndIf
        EndIf
        $oMail = Null
    Next

    If $bCategorized Then
        AdlibRegister("CategorizeEmails", 5000)
    Else
        AdlibRegister("CategorizeEmails", 40000)
    CleanupCOM()
EndFunc

Func HasPDFAttachment($oMail)
    Local $oAttachments = $oMail.Attachments
    For $i = 1 To $oAttachments.Count
        If StringInStr($oAttachments.Item($i).FileName, ".pdf") Then
            $oAttachments = Null
            Return True
        EndIf
    Next
    $oAttachments = Null
    Return False
EndFunc

Func IsRepliedOrForwarded($oMail)
    Local $sConversationIndex = $oMail.ConversationIndex
    Return StringLen($sConversationIndex) > 44
EndFunc

Func CleanupCOM()
    If IsObj($oCustomerScreeningFolder) Then
        $oCustomerScreeningFolder = Null
    EndIf
    If IsObj($oMailbox) Then
        $oMailbox = Null
    EndIf
    If IsObj($oNamespace) Then
        $oNamespace = Null
    EndIf
    If IsObj($oOutlook) Then
        $oOutlook = Null
    EndIf
EndFunc

Func WriteLog($sMessage)
    Local $sTimeStamp = _Now()
    Local $hFile = FileOpen($sLogFile, $FO_APPEND)
    If $hFile = -1 Then
        MsgBox($MB_SYSTEMMODAL, "Error", "Unable to open log file.")
        Return
    EndIf
    FileWriteLine($hFile, $sTimeStamp & " - " & $sMessage)
    FileClose($hFile)
EndFunc


CreateGUI()
