;***********************
;*** MORE CLIPBOARDS ***
;***********************

/*
;'''''''''''''''''''''''''''
;''' Available Functions '''
;'''''''''''''''''''''''''''

MoreClipboardsCopy(int) (1 - Clipboard Count)
MoreClipboardsPaste(int) (1 - Clipboard Count)
MoreClipboardsSend(int) (1 - Clipboard Count)
MoreClipboardsSendRaw(int) (1 - Clipboard Count)
MoreClipboardsGUI() (Launch Gui)
MoreClipboardsReloadVariables(FirstReloadVariable:=1)
MoreClipboardsPassParameters()

*/

MoreClipboards := []
MoreClipboardsDir := (A_AppData)"\AutoHotkey\MoreClipboards"
MoreClipboardsConfig := (MoreClipboardsDir)"\config.ini"
FileCreateDir, %MoreClipboardsDir%

; Default Clipboard Count.
; Change the loop below if more are desired
Loop, 10 
	MoreClipboards.Push(MoreClipboardsUtil_ReadClipboardFile(A_Index))

MoreClipboardsPasteMode_PASTE = 1
MoreClipboardsPasteMode_SEND = 2
MoreClipboardsPasteMode := MoreClipboardsUtil_GetSetting("PasteMode", MoreClipboardsPasteMode_PASTE)

MoreClipboardsCopy(ByRef index)
{
	Temp := ClipboardAll
	Clipboard := ""
	Send ^c
	ClipWait, 1
	ClipText := Clipboard
	Clipboard := Temp
	Temp := ""

	if ErrorLevel 
		return
	
	if WinActive("ahk_class XLMAIN")
		ClipText := RegExReplace(ClipText, "`r(?=`n)|`n$" "")
	
	global MoreClipboards
	MoreClipboards[index] := ClipText

	;CacheFile for later reload
	MoreClipboardsUtil_WriteClipboardFile(index, ClipText)
}

MoreClipboardsIndexToPasteOnChange := 0
MoreClipboardsPasteOnClipboardChange(Type) 
{
	global
	if (MoreClipboardsIndexToPasteOnChange) {
		Send, ^v
		Sleep 100
		MoreClipboardsIndexToPasteOnChange := 0
		Clipboard := ClipboardTemp
	}
}
OnClipboardChange("MoreClipboardsPasteOnClipboardChange")

MoreClipboardsPaste(ByRef index)
{
	global
	if MoreClipboardsPasteMode = %MoreClipboardsPasteMode_SEND%
	{
		Send, % MoreClipboards[index]
		return
	}

	ClipboardTemp := ClipboardAll
	MoreClipboardsIndexToPasteOnChange := 0
	Clipboard := ""
	MoreClipboardsIndexToPasteOnChange := index
	Clipboard := MoreClipboards[index]
}

MoreClipboardsOpenGUI()
{
	global
	local GuiTitle := "Set Clipboards Text"
	
	; This will stop the Gui window from resetting without saving your
	; clipboards if you press this hotkey and it's already open and instead give the window focus. 
	IfWinExist, %GuiTitle% ahk_class AutoHotkeyGUI 
	{
		WinActivate
		return
	}
	
	; If the window isn't open, it will create the Gui menu and display it.

	Gui, MoreClipboards:New,, %GuiTitle%
	Gui, Add, Text, % "x9 y6 w600", % "  Windows Clipboard"
	Gui, Add, Edit, % "vClipboard xp yp+15 wp r1", % Clipboard
	Loop, % MoreClipboards.Length()
	{
		Gui, Add, Text, % "xp yp+25 wp", % " Clipboard "(A_Index)
		Gui, Add, Edit, % "xp yp+15 wp vGuiMoreClipboards"(A_Index), % MoreClipboards[A_Index]
	}
	Gui, Add, Text, % "xp yp+35", Paste Mode
	Gui, Add, Radio, % "xp+65 yp vMoreClipboardsPasteMode "(MoreClipboardsPasteMode == MoreClipboardsPasteMode_PASTE ? "Checked" : ""), Paste
	Gui, Add, Radio, % "xp+50 yp "(MoreClipboardsPasteMode == MoreClipboardsPasteMode_SEND ? "Checked" : ""), Send
	Gui, Add, Button, % "x"(600-240)" yp-10 w120 h30", CLEAR`nCLIPBOARDS
	Gui, Add, Button, % "xp+130 yp wp hp Default", SET`nCLIPBOARDS
	Gui, Show
}

MoreClipboardsButtonSETCLIPBOARDS() 
{
	global
	Gui, Submit
	Loop, % MoreClipboards.Length() 
	{
		local content := GuiMoreClipboards%A_Index%
		GuiMoreClipboards%A_Index% := ""
		MoreClipboards[A_Index] := content
		MoreClipboardsUtil_WriteClipboardFile(A_Index, content)
	}

	MoreClipboardsUtil_SetSetting("PasteMode", MoreClipboardsPasteMode)
}

MoreClipboardsButtonCLEARCLIPBOARDS() 
{
	global
	Gui, Submit, NoHide
	Loop, % MoreClipboards.Length() 
	{
		GuiControl,,GuiMoreClipboards%A_Index%,
		MoreClipboards[A_Index] := ""
		MoreClipboardsUtil_WriteClipboardFile(A_Index, "")
	}
}

;Utility Functions not intended to be used outside of this Script

MoreClipboardsUtil_ReadClipboardFile(ByRef index)
{
	file := FileOpen(MoreClipboardsUtil_GetClipboardFilePath(index), "rw")		
	ClipText := file.Read()
	file.Close()
	return ClipText
}

MoreClipboardsUtil_WriteClipboardFile(ByRef index, ByRef content)
{
	global MoreClipboards
	file := FileOpen(MoreClipboardsUtil_GetClipboardFilePath(index), "w")		
	file.Write(MoreClipboards[index])
	file.Close()
}

MoreClipboardsUtil_GetClipboardFilePath(ByRef index)
{
	global MoreClipboardsDir
	return (MoreClipboardsDir)"\Clipboard"(index)
}

MoreClipboardsUtil_SetSetting(ByRef setting, ByRef value)
{
	global MoreClipboardsConfig
	IniWrite, %value%, %MoreClipboardsConfig%, "Settings", setting
}

MoreClipboardsUtil_GetSetting(ByRef setting, ByRef default)
{
	global MoreClipboardsConfig
	IniRead, value, %MoreClipboardsConfig%, "Settings", setting, %default%
	return value
}