;************************
;*** QUICK DATA ENTRY ***
;************************

/*
;'''''''''''''''''''''''''''
;''' Available Functions '''
;'''''''''''''''''''''''''''

QuickDataEntry_SendNextQuickText()
QuickDataEntry_GUI()
QuickDataEntry_StartOver()
QuickDataEntry_StepBack()

;'''''''''''''''''''''''''''''''''''''
;''' Default Hotkey Configurations '''
;'''''''''''''''''''''''''''''''''''''

Insert::QuickDataEntry_SendNextQuickText()
!Insert::QuickDataEntry_GUI()
^Insert::QuickDataEntry_StartOver()
+Insert::QuickDataEntry_StepBack()

*/

/*
;Test
!Insert::
PRIVATE_QuickDataEntry_StringToArray("A~B~C~D~E","~")
QuickDataEntry_ParseNum = 1
QuickDataEntry_LoopNum = 1
return
;
*/

QuickDataEntry_SendNextQuickText()
{
	global
	if not QuickDataEntry_LoopNum
		return

	loop, %QuickDataEntry_LoopNum%
	{
		;Send the data at the Specified Index
		Send % QuickDataEntry_QuickArray[QuickDataEntry_ParseNum] ;Lower Bound of Array is 1
		
		QuickDataEntry_ParseNum++
		if QuickDataEntry_ParseNum > % QuickDataEntry_QuickArray.MaxIndex()
			{
				QuickDataEntry_ParseNum = 0
				QuickDataEntry_LoopNum = 0
				break
			}
	}
}

QuickDataEntry_GUI()
{
	global
	local wGui = 700
	local wText = 80
	local wButton = 100
	if QuickDataEntry_ParseNum = 0
		QuickDataEntry_QuickText := ""
	else
		QuickDataEntry_QuickText := QuickDataEntry_QuickArray[QuickDataEntry_ParseNum]
	
	Gui, QuickDataEntry:New
	Gui, +AlwaysOnTop
	Gui, Add	, Text			, % "Center 			x0 										y15 w"(wGui/2)	, Parse Position
	Gui, Add	, Edit			, % "vQuickDataEntry_ParseNum Center 	x"(((wGui - wText)/2)-(wGui/4))" 		y33 w"(wText)	, %QuickDataEntry_ParseNum%
	Gui, Add	, Text			, % "Center 			x"(wGui/2)" 							y15 w"(wGui/2)	, Number of Loops
	Gui, Add	, Edit			, % "vQuickDataEntry_LoopNum Center 	x"(((wGui - wText)/2)+(wGui/4))" 		y33 w"(wText)	, %QuickDataEntry_LoopNum%
	Gui, Add	, Button		, % "Default 			x"((wGui/4) - (wButton/2))"	y60 w"(wButton)	, Current &Data Set
	Gui, Add	, Button		, % "					x"((wGui/4) - (wButton/2) + (wGui/2))" 			y60 w"(wButton)	, &Clipboard Set
	Gui, Add	, Text			, % "				 	x14  w"(wGui-18) , Next Quick Text
	Gui, Add	, Edit			, % "vQuickDataEntry_QuickText x12 y105 w"(wGui - 25), %QuickDataEntry_QuickText%
	Gui, Show	, % " w"(wGui)	, Quick Data Entry
}
/*
Gui, Add, Text, x22 y20 w90 h30 , Parse Position
Gui, Add, Text, x22 y60 w90 h30 , Number of Loops
Gui, Add, Button, x372 y20 w100 h30 , Button2
Gui, Add, Button, x262 y20 w100 h30 , Button1
Gui, Add, Edit, x122 y20 w100 h30 , Edit
Gui, Add, Edit, x122 y60 w100 h30 , Edit
Gui, Add, Edit, x12 y190 w550 h20 , Edit
Gui, Add, Text, x12 y160 w130 h20 , Text
Gui, Add, ListBox, x12 y220 w550 h540 , ListBox
Gui, Add, Button, x262 y60 w100 h30 , Button3
Gui, Add, Button, x372 y60 w100 h30 , Button4
Gui, Add, Edit, x122 y100 w100 h30 , Edit
Gui, Add, Text, x22 y100 w90 h30 , Delimiter
; Generated using SmartGUI Creator for SciTE
Gui, Show, w600 h800, Untitled GUI
return

GuiClose:
ExitApp
*/

QuickDataEntryButtonCURRENTDATASET:
	Gui, QuickDataEntry:Submit
	PRIVATE_QuickDataEntry_EditQuickArrayElement()
	return

QuickDataEntryButtonCLIPBOARDSET:
	PRIVATE_QuickDataEntry_StringToArray(Clipboard)
	Gui, QuickDataEntry:Submit
	return

QuickDataEntry_StartOver()
{
	global
	QuickDataEntry_ParseNum = 1
	if Not QuickDataEntry_LoopNum
		QuickDataEntry_LoopNum = 1
}

QuickDataEntry_StepBack()
{
	global
	
	if (QuickDataEntry_ParseNum - QuickDataEntry_LoopNum) > 1
		QuickDataEntry_ParseNum -= QuickDataEntry_LoopNum
	else
		QuickDataEntry_ParseNum = 1
}

PRIVATE_QuickDataEntry_StringToArray(NewQuickString,sDelimiter:="~")
{
	global QuickDataEntry_QuickArray
	QuickDataEntry_QuickArray := StrSplit(NewQuickString,sDelimiter) ;The Lowerbound of this Array is 1
}

PRIVATE_QuickDataEntry_EditQuickArrayElement()
{
	global
	QuickDataEntry_QuickArray[QuickDataEntry_ParseNum] := QuickDataEntry_QuickText
}

/*
PRIVATE_QuickDataEntry_ExcelToArray(NewQuickString,sDelimiter:="~")
{
	global QuickDataEntry_QuickArray
	for cell in 
	QuickDataEntry_QuickArray := StrSplit(NewQuickString,sDelimiter) ;The Lowerbound of this Array is 1
}
*/
	
/*
;Microsoft Excel Window
#IfWinActive ahk_class XLMAIN
^Insert::
QuickDataEntry_ParseNum = 1
if Not QuickDataEntry_LoopNum
	QuickDataEntry_LoopNum = 1

QuickText := ComObjActive("Excel.Application").Selection
MsgBox %QuickText%
StringReplace QuickText, QuickText, `r`n, ~
MsgBox %QuickText%

return
#IfWinActive
*/