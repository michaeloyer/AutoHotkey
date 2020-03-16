;************************
;*** QUICK DATA ENTRY ***
;************************

/*
;'''''''''''''''''''''''''''
;''' Available Functions '''
;'''''''''''''''''''''''''''

QuickDataEntry_SendNextQuickText()
QuickDataEntry_GUI()

*/

QuickDataEntry_SendNextQuickText()
{
	global

	if (QuickDataEntry_ArrayIndex == -1)
		return

	if (QuickDataEntry_ArrayIndex > QuickDataEntry_QuickArray.MaxIndex())
	{
		QuickDataEntry_ArrayIndex := -1
		return
	}

	if (QuickDataEntry_LoopNum == 0) 
	{
		for index, element in QuickDataEntry_QuickArray 
		{
			if (index < QuickDataEntry_ArrayIndex)
				continue
			
			Send, %element%
			QuickDataEntry_ArrayIndex++
		}
	}
	else
	{
		Send % QuickDataEntry_QuickArray[QuickDataEntry_ArrayIndex]
		QuickDataEntry_ArrayIndex++
	}
}

QuickDataEntry_GUI()
{
	global

	local GuiTitle := "Quick Data Entry"
	
	; This will stop the Gui window from resetting without saving your
	; clipboards if you press this hotkey and it's already open and instead give the window focus. 
	IfWinExist, %GuiTitle% ahk_class AutoHotkeyGUI 
	{
		WinActivate
		return
	}

	if not QuickDataEntry_LoopNum
		QuickDataEntry_LoopNum := 1

	if not QuickDataEntry_DelimiterSelection
		QuickDataEntry_DelimiterSelection := 1

	Gui, QuickDataEntry:New,, %GuiTitle%
	Gui, Add, GroupBox, x12 w160 h100, Delimiter
	Gui, Add, Radio, % "x20 yp+20 vQuickDataEntry_DelimiterSelection "(QuickDataEntry_DelimiterSelection == 1 ? "Checked" : "")" section", CRLF
	Gui, Add, Radio, % (QuickDataEntry_DelimiterSelection == 2 ? "Checked" : ""), LF
	Gui, Add, Radio, % (QuickDataEntry_DelimiterSelection == 3 ? "Checked" : ""), TAB
	Gui, Add, Radio, % (QuickDataEntry_DelimiterSelection == 4 ? "Checked" : ""), SPACE
	Gui, Add, Radio, % "ys "(QuickDataEntry_DelimiterSelection == 5 ? "Checked" : ""), Other
	Gui, Add, Edit, % "vQuickDataEntry_DelimiterOther Center w80", % (QuickDataEntry_DelimiterSelection == 5 ? QuickDataEntry_DelimiterOther : "")

	Gui, Add, Text, x180 y15, Number of Loops
	Gui, Add, Edit, % "Center Number xp+90 yp-3 w70", %QuickDataEntry_LoopNum%
	Gui, Add, UpDown, Range0-32767 vQuickDataEntry_LoopNum, %QuickDataEntry_LoopNum%
	
	Gui, Add, Text, x12, Quick Text
	Gui, Add, Edit, % "vQuickDataEntry_QuickText x12 yp+13 w610 r20", %QuickDataEntry_QuickText%
	
	Gui, Add, Button, % "w70 section", Set Text
	Gui, Add, Button, % "ys w50", OK
	Gui, Add, Button, % "ys w50", Cancel
	
	Gui +Delimiter`b,
	Gui, Add, Text, x630 y10, Quick Array Items
	Gui, Add, ListBox, xp yp+15 w360 r28 vQuickDataEntry_ArrayIndex AltSubmit, % PRIVATE_QuickDataEntry_GetListBoxValues()
	
	Gui, Show
}

;GUI Functions
QuickDataEntryButtonSETTEXT()
{
	Gui, QuickDataEntry:Submit, NoHide
	PRIVATE_QuickDataEntry_SetQuickData()
}

QuickDataEntryButtonCANCEL()
{
	Gui, QuickDataEntry:Cancel
}

QuickDataEntryButtonOK()
{
	Gui, QuickDataEntry:Submit
	PRIVATE_QuickDataEntry_SetQuickData()
}

;Private Functions
PRIVATE_QuickDataEntry_SetQuickData()
{
	global 
	PRIVATE_QuickDataEntry_SetDelimiter() 
	PRIVATE_QuickDataEntry_SetQuickArray()

	if not QuickDataEntry_ArrayIndex
		QuickDataEntry_ArrayIndex := 1

	if (QuickDataEntry_QuickArray.Length() < QuickDataEntry_ArrayIndex)
		QuickDataEntry_ArrayIndex := 1

	GuiControl, Text, QuickDataEntry_QuickText, %QuickDataEntry_QuickText%
	
	;Empty before appending new list items
	GuiControl,, QuickDataEntry_ArrayIndex, `b
	GuiControl,, QuickDataEntry_ArrayIndex, % PRIVATE_QuickDataEntry_GetListBoxValues()
}

PRIVATE_QuickDataEntry_SetDelimiter() 
{
	global
	QuickDataEntry_Delimiter := (QuickDataEntry_DelimiterSelection) == 1 
			? "`r`n"
		: (QuickDataEntry_DelimiterSelection) == 2 
			? "`n"
		: (QuickDataEntry_DelimiterSelection) == 3 
			? "`t"
		: (QuickDataEntry_DelimiterSelection) == 4 
			? " "
		: (QuickDataEntry_DelimiterSelection) == 5 
			? QuickDataEntry_DelimiterOther
		: ""
}

PRIVATE_QuickDataEntry_SetQuickArray() 
{
	global 
	QuickDataEntry_QuickArray := []
	text := ""

	local array := StrSplit(QuickDataEntry_QuickText, QuickDataEntry_Delimiter)

	local filteredArray := []
	for index, element in StrSplit(QuickDataEntry_QuickText, QuickDataEntry_Delimiter) 
	{
		if (element != "")
			filteredArray.push(element)
	}
	array := filteredArray
	
	for index, element in array
	{
		text .= element
		if (Mod(index, QuickDataEntry_LoopNum) == 0) 
		{
			QuickDataEntry_QuickArray.Push(text)
			text := ""
		}
	}

	if (text)
		QuickDataEntry_QuickArray.Push(text)
	
}

PRIVATE_QuickDataEntry_GetListBoxValues() 
{
	global 
	local ListBoxValues := ""
	for index, element in QuickDataEntry_QuickArray 
	{
		ListBoxValues .= element
		ListBoxValues .= "`b"

		if (index == QuickDataEntry_ArrayIndex)
			ListBoxValues .= "`b"
	}
	return ListBoxValues
}
