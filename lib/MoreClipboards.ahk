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
FileCreateDir, %MoreClipboardsDir%

; Default Clipboard Count.
; Change the loop below if more are desired
Loop, 10 
{
	MoreClipboards.Push(MoreClipboardsUtil_ReadClipboardFile(A_Index))
}

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
	ClipboardTemp := ClipboardAll
	MoreClipboardsIndexToPasteOnChange := 0
	Clipboard := ""
	MoreClipboardsIndexToPasteOnChange := index
	Clipboard := MoreClipboards[index]
}

MoreClipboardsSend(ByRef index)
{
	global
	Send % MoreClipboards[index]
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
	wGui := 600,  wButton := wGui/5
	Gui, MoreClipboards:New,, %GuiTitle%
	Gui, Add, Text, %           " x9 y6    w"(wGui), % "  Windows Clipboard"
	Gui, Add, Edit, % "vClipboard xp yp+15 wp r1", % Clipboard
	Loop, % MoreClipboards.Length()
	{
		Gui, Add, Text, % "                              xp yp+25 wp", % " Clipboard "(A_Index)
		Gui, Add, Edit, % "vGuiMoreClipboards"(A_Index)" xp yp+15 wp", % MoreClipboards[A_Index]
	}
	Gui, Add, Button, % " Default x"(9+(wGui/2)-(wButton+10))" yp+25 w"(wButton)" h30", SET CLIPBOARDS
	Gui, Add, Button, % "         x"(9+(wGui/2)+10)"           yp    wp           hp ", CLEAR
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
}

MoreClipboardsButtonCLEAR() 
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

MoreClipboardsUtil_CleanParameterPassingString(PassString) ;Backslashes Nullify Quotes and other Backslashes.  See the help file for more details.
{
	If InStr(PassString,"""") > 0 ;If the parameter being passed has a Quotation Mark or Backslash, the string will have to be put through these series of loops.
	{		
		SearchString := """"
		;This while loop will find the maximum number of backslashes followed by one quotation mark.
		While % InStr(PassString,"\" . SearchString) > 0
			SearchString := "\" . SearchString
		
		;This loop adds a backslash to a series of backslashes followed by a quotation mark.
		;By the end of the loop all back slashes in the 
		Loop, % StrLen(SearchString)
		{			
			StringReplace, PassString, PassString, % SearchString, % "\" . SearchString, All
			SearchString := SubStr(SearchString,2) ;Take Backslashes away one by one.
		}
	}
	If Instr(PassString,"\") > 0 
	{
		Counter := 0
		;Strings that have ending backslashes (one or more) will be taken care of in this while loop
		While % SubStr(PassString, -Counter, 1) = "\" And StrLen(PassString) > Counter
			Counter++
		
		PassString .= SubStr(PassString, - (Counter - 1), Counter)
	}
	
	PassString := """" . PassString . """" ;Put 1 set of quotation marks around the PassString.
	Return PassString
}

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