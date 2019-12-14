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
	TenClipboardsReloadVariables(1)
	
#Include CaseText.ahk
#Include Script Building.ahk
#Include Ten Clipboards.ahk
#Include Quick Data Entry.ahk

Pause::Suspend
NumLock & NumpadAdd::Send {Volume_Up}
NumLock & NumpadSub::Send {Volume_Down}
NumLock & NumpadMult::Send {Volume_Mute}

#Hotstring ? o1 
#Hotstring EndChars `t`n


ScrollLock & NumLock::Reload
NumLock & ScrollLock::Run, % """" . A_AhkPath . """ /restart """ . A_ScriptFullPath . """ " . TenClipboardsPassParameters()

;Open the Clipboard GUI
Alt & `::TenClipboardsOpenGUI()

;Copy to the Ten Clipboards (Ctrl + Number)
^1:: TenClipboardsCopy(1)
^2:: TenClipboardsCopy(2)
^3:: TenClipboardsCopy(3)
^4:: TenClipboardsCopy(4)
^5:: TenClipboardsCopy(5)
^6:: TenClipboardsCopy(6)
^7:: TenClipboardsCopy(7)
^8:: TenClipboardsCopy(8)
^9:: TenClipboardsCopy(9)
^0:: TenClipboardsCopy(0)

;Send Text from the Ten Clipboards (Alt + Number)
!1:: TenClipboardsSend(1)
!2:: TenClipboardsSend(2)
!3:: TenClipboardsSend(3)
!4:: TenClipboardsSend(4)
!5:: TenClipboardsSend(5)
!6:: TenClipboardsSend(6)
!7:: TenClipboardsSend(7)
!8:: TenClipboardsSend(8)
!9:: TenClipboardsSend(9)
!0:: TenClipboardsSend(0)

;Send Literal Text from the Clipboards (Shift + Alt + Number)
+!1:: TenClipboardsSendRaw(1)
+!2:: TenClipboardsSendRaw(2)
+!3:: TenClipboardsSendRaw(3)
+!4:: TenClipboardsSendRaw(4)
+!5:: TenClipboardsSendRaw(5)
+!6:: TenClipboardsSendRaw(6)
+!7:: TenClipboardsSendRaw(7)
+!8:: TenClipboardsSendRaw(8)
+!9:: TenClipboardsSendRaw(9)
+!0:: TenClipboardsSendRaw(0)

Insert::QuickDataEntry_SendNextQuickText()
Numlock & Insert::QuickDataEntry_GUI()
^Insert::QuickDataEntry_StartOver()
+Insert::QuickDataEntry_StepBack()