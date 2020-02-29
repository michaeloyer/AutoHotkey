;**********************
;*** TEN CLIPBOARDS ***
;**********************

/*
;'''''''''''''''''''''''''''
;''' Available Functions '''
;'''''''''''''''''''''''''''

TenClipboardsCopy(int) (1 - Clipboard Count)
TenClipboardsPaste(int) (1 - Clipboard Count)
TenClipboardsSend(int) (1 - Clipboard Count)
TenClipboardsSendRaw(int) (1 - Clipboard Count)
TenClipboardsGUI() (Launch Gui)
TenClipboardsReloadVariables(FirstReloadVariable:=1)
TenClipboardsPassParameters()


;''''''''''''''''''''''''''''''''''''''''''''''''''''
;''' Restarting the Script with Clipboards Intact '''
;''''''''''''''''''''''''''''''''''''''''''''''''''''

To restart the script with all clipboards values intact you can use the following line in place of Reload:

Run, % """" . A_AhkPath . """ /restart """ . A_ScriptFullPath . """ " . TenClipboardsPassParameters()


This should go in your Auto-Exectue Section:

;Preserve 10 Clipboards variables if script gets restarted as opposed to reloaded 
If %0% <> 0
	TenClipboardsReloadVariables(1)

;'''''''''''''''''''''''''''''''''''''
;''' Default Hotkey Configurations '''
;'''''''''''''''''''''''''''''''''''''

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
^0:: TenClipboardsCopy(10)

;Paste Text from the Ten Clipboards (Alt + Number)
!1:: TenClipboardsPaste(1)
!2:: TenClipboardsPaste(2)
!3:: TenClipboardsPaste(3)
!4:: TenClipboardsPaste(4)
!5:: TenClipboardsPaste(5)
!6:: TenClipboardsPaste(6)
!7:: TenClipboardsPaste(7)
!8:: TenClipboardsPaste(8)
!9:: TenClipboardsPaste(9)
!0:: TenClipboardsPaste(10)

;Send Text from the Ten Clipboards (Ctrl + Alt + Number)
^!1:: TenClipboardsSend(1)
^!2:: TenClipboardsSend(2)
^!3:: TenClipboardsSend(3)
^!4:: TenClipboardsSend(4)
^!5:: TenClipboardsSend(5)
^!6:: TenClipboardsSend(6)
^!7:: TenClipboardsSend(7)
^!8:: TenClipboardsSend(8)
^!9:: TenClipboardsSend(9)
^!0:: TenClipboardsSend(10)

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
+!0:: TenClipboardsSendRaw(10)

*/

TenClipboards := []
Loop, 10 ; Default Clipboard Count
	TenClipboards.Push("")

TenClipboardsCopy(ByRef index)
{
	global
	local ClipboardTemp := ClipboardAll
	Clipboard := ""
	Send ^c
	ClipWait, 1
	
	if not ErrorLevel
	{
		if WinActive("ahk_class XLMAIN")
			PRIVATE_TenClipboards_Excel_To_Clipboard()
		
		TenClipboards[index] := Clipboard		
	}

	Clipboard := ClipboardTemp
	ClipboardTemp := ""
}

TenClipboardsIndexToPasteOnChange := 0
TenClipboardsPasteOnClipboardChange(Type) 
{
	global
	if (TenClipboardsIndexToPasteOnChange) {
		Send, ^v
		Sleep 25
		TenClipboardsIndexToPasteOnChange := 0
		Clipboard := ClipboardTemp
	}
}
OnClipboardChange("TenClipboardsPasteOnClipboardChange")


TenClipboardsPaste(ByRef index)
{
	global
	ClipboardTemp := ClipboardAll
	TenClipboardsIndexToPasteOnChange := 0
	Clipboard := ""
	TenClipboardsIndexToPasteOnChange := index
	Clipboard := TenClipboards[index]
}

TenClipboardsSend(ByRef index)
{
	global
	Send % TenClipboards[index]
}

TenClipboardsSendRaw(ByRef index)
{
	global
	SendRaw % TenClipboards[index]
}

TenClipboardsOpenGUI()
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
	
	xGui := 9,  yGui := 0,  wGui := 600,  hGui := 20,  wButton := wGui/5,  ySection := 40
	Gui, TenClipboards:New,, %GuiTitle%
	Gui, Add, Text, %           " x"(xGui)" y"( 6+yGui)" w"(wGui) " h"(hGui) , % "  Windows Clipboard"
	Gui, Add, Edit, % "vClipboard x"(xGui)" y"(21+yGui)" w"(wGui) " r3", % Clipboard
	Loop, % TenClipboards.Length()
	{
		yGui += ySection
		Gui, Add, Text, %                             " x"(xGui)" y"( 6+yGui)" w"(wGui) " h"(hGui) , % " Clipboard "(A_Index)
		Gui, Add, Edit, % "vGuiTenClipboards"(A_Index)" x"(xGui)" y"(21+yGui)" w"(wGui) " h"(hGui) , % TenClipboards[A_Index]
	}
	yGui += ySection
	Gui, Add, Button, % " Default x"(9+(wGui/2)-(wButton/2))" y"(6+yGui)" w"(wButton)" h30", SET CLIPBOARDS
	Gui, Show
}

TenClipboardsButtonSETCLIPBOARDS:
	Gui, Submit
	Loop, % TenClipboards.Length() 
	{
		TenClipboards[A_Index] := GuiTenClipboards%A_Index%
		GuiTenClipboards%A_Index% := ""
	}
	return

/*
Used for resetting a script without having to lose your 10 Clipboards values
even when a clipboard is empty it will pass an empty string, so 10 variables
are passed into the reloaded script.

TenClipboardsReloadVariables should be in your main script's auto-execute section.

For more information see "parameters passed into a script" in the AutoHotkey Help Index or search for 
"Passing Command Line Parameters to a Script".
*/

;Change FirstReloadVariable if you are passing other variables first ahead of the 10 Clip# variables
TenClipboardsReloadVariables(FirstReloadVariable:=1) 
{
	global
	Loop, % TenClipboards.Length()
	{
		PassParamNum := A_Index + FirstReloadVariable - 1
		TenClipboards[%A_LoopField%] := %PassParamNum%
	}
}

TenClipboardsPassParameters()
{
	global
	local SingleClipboard := ""

	Loop, % TenClipboards.Length()
	{
		SingleClipboard := TenClipboards[A_Index]
		PassParameters .= " " . PRIVATE_TenClipboards_CleanParameterPassingString(%SingleClipboard%)
	}

	return PassParameters
}

;Private Functions not intended to be used outside of this Script

PRIVATE_TenClipboards_Excel_To_Clipboard()
{
	global
	StringReplace, Clipboard, Clipboard, `r`n , `n, All
	if SubStr(Clipboard,0,1) = Chr(10) ;Linefeed Character
		StringLeft, Clipboard, Clipboard, StrLen(Clipboard) - 1
}

PRIVATE_TenClipboards_CleanParameterPassingString(PassString) ;Backslashes Nullify Quotes and other Backslashes.  See the help file for more details.
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