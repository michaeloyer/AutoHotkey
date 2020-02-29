;******************
;*** Directives ***
;******************

#SingleInstance Force
#NoEnv

SendMode Input
SetWorkingDir %A_ScriptDir%
SetNumLockState, AlwaysOn
SetScrollLockState, AlwaysOff
SetTitleMatchMode, 2

;*************************
;*** Auto Launch Field ***
;*************************

#Include <MoreClipboards>
#Include <CaseText>
#Include <QuickDataEntry>

#Include %A_ScriptDir%\Hotkeys
#Include Script Building.ahk

Pause::Suspend
NumLock & NumpadAdd::Send {Volume_Up}
NumLock & NumpadSub::Send {Volume_Down}
NumLock & NumpadMult::Send {Volume_Mute}

#Hotstring ? o1 
#Hotstring EndChars `t`n

ScrollLock & NumLock::Reload
NumLock & ScrollLock::Reload

;Open the Clipboard GUI
Alt & `::MoreClipboardsOpenGUI()

;Copy to Clipboards (Ctrl + Number)
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

;Paste from Clipboards (Alt + Number)
!1:: MoreClipboardsPaste(1)
!2:: MoreClipboardsPaste(2)
!3:: MoreClipboardsPaste(3)
!4:: MoreClipboardsPaste(4)
!5:: MoreClipboardsPaste(5)
!6:: MoreClipboardsPaste(6)
!7:: MoreClipboardsPaste(7)
!8:: MoreClipboardsPaste(8)
!9:: MoreClipboardsPaste(9)
!0:: MoreClipboardsPaste(10)

;Send Text from Clipboards (Shift + Alt + Number)
+!1:: MoreClipboardsSend(1)
+!2:: MoreClipboardsSend(2)
+!3:: MoreClipboardsSend(3)
+!4:: MoreClipboardsSend(4)
+!5:: MoreClipboardsSend(5)
+!6:: MoreClipboardsSend(6)
+!7:: MoreClipboardsSend(7)
+!8:: MoreClipboardsSend(8)
+!9:: MoreClipboardsSend(9)
+!0:: MoreClipboardsSend(10)

Insert::QuickDataEntry_SendNextQuickText()
!Insert::QuickDataEntry_GUI()
^Insert::QuickDataEntry_StartOver()
+Insert::QuickDataEntry_StepBack()

#CapsLock::CapsLock
CapsLock::CaseTextNext()
+CapsLock::CaseTextUpper()
!CapsLock::CaseTextLower()
^CapsLock::CaseTextTitle()