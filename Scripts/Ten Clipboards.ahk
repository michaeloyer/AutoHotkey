;**********************
;*** TEN CLIPBOARDS ***
;**********************

/*
;'''''''''''''''''''''''''''
;''' Available Functions '''
;'''''''''''''''''''''''''''

TenClipboardsCopy(int) (Must be 0 - 9)
TenClipboardsPaste(int) (Must be 0 - 9)
TenClipboardsSend(int) (Must be 0 - 9)
TenClipboardsSendRaw(int) (Must be 0 - 9)
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
^0:: TenClipboardsCopy(0)

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
!0:: TenClipboardsPaste(0)

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
^!0:: TenClipboardsSend(0)

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

*/



TenClipboardsCopy(ByRef ButtonIndex)
{
	global
	local ClipboardTemp := Clipboard
	Clipboard := ""
	Send ^c
	ClipWait, 1
	
	if not ErrorLevel
	{
		if WinActive("ahk_class XLMAIN")
			PRIVATE_TenClipboards_Excel_To_Clipboard()
		
		TenClipboards_Clip%ButtonIndex% := Clipboard		
	}

	Clipboard := ClipboardTemp
	ClipboardTemp := ""
}

TenClipboardsToPasteOnChange := 0
TenClipboardsPasteOnClipboardChange(Type) 
{
	global
	if (TenClipboardsToPasteOnChange) {
		Send, ^v
		Sleep 25
		TenClipboardsToPasteOnChange := 0
		Clipboard := ClipboardTemp
	}
}
OnClipboardChange("TenClipboardsPasteOnClipboardChange")


TenClipboardsPaste(ByRef ButtonIndex)
{
	global
	ClipboardTemp := ClipboardAll
	TenClipboardsToPasteOnChange := 0
	Clipboard := ""
	TenClipboardsToPasteOnChange := ButtonIndex
	Clipboard := TenClipboards_Clip%ButtonIndex%
}

TenClipboardsSend(ByRef ButtonIndex)
{
	global
	Send % TenClipboards_Clip%ButtonIndex%
}

TenClipboardsSendRaw(ByRef ButtonIndex)
{
	global
	SendRaw % TenClipboards_Clip%ButtonIndex%
}

TenClipboardsOpenGUI()
{
	global
	local ClipboardLoops := 10
	local GuiTitle := "Set Clipboards Text"
	  ; This will stop the Gui window from resetting without saving your
	  ; clipboards if you press this hotkey and it's already open and instead give the window focus. 
	IfWinExist, %GuiTitle% ahk_class AutoHotkeyGUI
		WinActivate
	else
	{
	  ; If the window isn't open, it will create the Gui menu and display it.
		Gui, TenClipboards:New,, %GuiTitle%
		xGui := 9,   yGui := 0,   wGui := 600,   hGui := 20,   wButton := wGui/5
		Gui, Add, Text, %           " x"(xGui)" y"( 6+yGui)" w"(wGui) " h"(hGui) , % "  Windows Clipboard"
		Gui, Add, Edit, % "vClipboard x"(xGui)" y"(21+yGui)" w"(wGui) " h"(hGui) " r1", % Clipboard
		Loop, %ClipboardLoops%
		{
			yGui := ((A_Index)*40)
			ButtonIndex := A_Index
			ButtonIndex := A_Index=10 ? "0" : ButtonIndex
			Gui, Add, Text, %                     " x"(xGui)" y"( 6+yGui)" w"(wGui) " h"(hGui) , % " Clipboard "(ButtonIndex)
			Gui, Add, Edit, % "vTenClipboards_Clip"(ButtonIndex)" x"(xGui)" y"(21+yGui)" w"(wGui) " h"(hGui) , %   TenClipboards_Clip%ButtonIndex%
		}
		Gui, Add, Button, % " Default x"(9+(wGui/2)-(wButton/2))" y"(6+((ClipboardLoops+1)*40))" w"(wButton)" h30", SET CLIPBOARDS
		Gui, Show
	}
}

TenClipboardsButtonSETCLIPBOARDS:
	Gui, Submit
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
	local Clips := 1234567890
	local PassParamNum := 0
	Loop, Parse, Clips
	{
		PassParamNum := A_Index + FirstReloadVariable - 1
		TenClipboards_Clip%A_LoopField% := %PassParamNum%
	}
}

TenClipboardsPassParameters()
{
	global
	local PassParameters := ""
	local Clips := 1234567890
	local SingleClipboard := ""

	Loop, Parse, Clips
	{
		SingleClipboard := "TenClipboards_Clip" . A_LoopField
		PassParameters .= " " . PRIVATE_TenClipboards_CleanParameterPassingString(%SingleClipboard%)
	}
	Return PassParameters
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