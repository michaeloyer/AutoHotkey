;******************
;*** Directives ***
;******************

#SingleInstance Force
#NoEnv
#Warn UseUnsetGlobal, Off

SendMode Input
SetWorkingDir %A_ScriptDir%
SetNumLockState, AlwaysOn
SetScrollLockState, AlwaysOff
SetTitleMatchMode, 2

;*************************
;*** Auto Launch Field ***
;*************************

;Preserve 10 Clipboards variables if script gets restarted as opposed to reloaded
If %0% <> 0
	MoreClipboardsReloadVariables(1)

#Include <MoreClipboards>

#Include %A_ScriptDir%\Scripts
#Include CaseText.ahk
#Include QuickDataEntry.ahk

#Include %A_ScriptDir%\Hotkeys
#Include Script Building.ahk

Pause::Suspend
NumLock & NumpadAdd::Send {Volume_Up}
NumLock & NumpadSub::Send {Volume_Down}
NumLock & NumpadMult::Send {Volume_Mute}

#Hotstring ? o1 
#Hotstring EndChars `t`n

ScrollLock & NumLock::Reload
NumLock & ScrollLock::Run, % """" . A_AhkPath . """ /restart """ . A_ScriptFullPath . """ " . MoreClipboardsPassParameters()

;Open the Clipboard GUI
Alt & `::MoreClipboardsOpenGUI()

;Copy to the Ten Clipboards (Ctrl + Number)
^1:: MoreClipboardsCopy(1)
^2:: MoreClipboardsCopy(2)
^3:: MoreClipboardsCopy(3)
^4:: MoreClipboardsCopy(4)
^5:: MoreClipboardsCopy(5)
^6:: MoreClipboardsCopy(6)
^7:: MoreClipboardsCopy(7)
^8:: MoreClipboardsCopy(8)
^9:: MoreClipboardsCopy(9)
^0:: MoreClipboardsCopy(10)

;Send Text from the Ten Clipboards (Alt + Number)
!1:: MoreClipboardsSend(1)
!2:: MoreClipboardsSend(2)
!3:: MoreClipboardsSend(3)
!4:: MoreClipboardsSend(4)
!5:: MoreClipboardsSend(5)
!6:: MoreClipboardsSend(6)
!7:: MoreClipboardsSend(7)
!8:: MoreClipboardsSend(8)
!9:: MoreClipboardsSend(9)
!0:: MoreClipboardsSend(10)

;Send Literal Text from the Clipboards (Shift + Alt + Number)
+!1:: MoreClipboardsSendRaw(1)
+!2:: MoreClipboardsSendRaw(2)
+!3:: MoreClipboardsSendRaw(3)
+!4:: MoreClipboardsSendRaw(4)
+!5:: MoreClipboardsSendRaw(5)
+!6:: MoreClipboardsSendRaw(6)
+!7:: MoreClipboardsSendRaw(7)
+!8:: MoreClipboardsSendRaw(8)
+!9:: MoreClipboardsSendRaw(9)
+!0:: MoreClipboardsSendRaw(10)

Insert::QuickDataEntry_SendNextQuickText()
Numlock & Insert::QuickDataEntry_GUI()
^Insert::QuickDataEntry_StartOver()
+Insert::QuickDataEntry_StepBack()