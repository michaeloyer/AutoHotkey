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

;*******************************
;*** Auto Excecution Section ***
;*******************************

#Include <MoreClipboards>
#Include <CaseText>
#Include <QuickDataEntry>
#Include *i %A_MyDocuments%/AutoHotkey/Functions.ahk

FileCreateDir, % (A_MyDocuments)"/AutoHotkey"
FileAppend,, %A_MyDocuments%\AutoHotkey\Hotkeys.ahk
FileAppend,, %A_MyDocuments%\AutoHotkey\Functions.ahk

;***************
;*** Hotkeys ***
;***************

#Include <HotkeyWriterHotkeys>
#Include <MoreClipboardsHotkeys>
#Include <CaseTextHotkeys>
#Include <QuickDataEntryHotkeys>
#Include *i %A_MyDocuments%/AutoHotkey/Hotkeys.ahk

Pause::Suspend
NumLock & NumpadAdd::Send {Volume_Up}
NumLock & NumpadSub::Send {Volume_Down}
NumLock & NumpadMult::Send {Volume_Mute}

EditUserAhkFile(ahkFilePath)
{
    workingDirectory := (A_MyDocuments)"\AutoHotkey"
    ;Open in Visual Studio Code
    Run, % "code.cmd "(ahkFilePath), %workingDirectory%, Hide UseErrorLevel
    if ErrorLevel
        ;Open in Notepad
        Run, % "notepad.exe "(ahkFilePath), %workingDirectory%
}

!#e::EditUserAhkFile(A_ScriptFullPath)
!#h::EditUserAhkFile("Hotkeys.ahk")
!#f::EditUserAhkFile("Functions.ahk")

#Hotstring ? o1 
#Hotstring EndChars `t`n

ScrollLock & NumLock::Reload
NumLock & ScrollLock::Reload
