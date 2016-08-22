;****************
;*** CaseText ***
;****************

/*
This Script is used to capitalized highlighted strings via
The Windows Clipboard.  The 4 Case options available are:

CaseText("Next")
CaseText("Upper")
CaseText("Lower")
CaseText("Title")

*/

CaseText(CapsOption)
{
	ClipboardTemp := ""
	ClipboardTemp := ClipboardAll
	Clipboard := ""
	
	Send ^c 
	ClipWait, 1
	If Not Clipboard 
		Return
	Clipboard := Clipboard ;Converting clipboard to plain text
	sString := Clipboard ;Convert clipboard contents to a variable
	Clipboard := ClipboardTemp ;Return what was previously on the clipboard 

		;The 'Next' CapsOption will let you cycle the case.
		If CapsOption = Next
		{
			sAlphaString := PRIVATE_CaseText_TextOnly(sString)
			If sAlphaString is upper
				StringLower, sString, sString
			Else If sAlphaString is lower
				StringUpper, sString, sString, T
			Else
				StringUpper, sString, sString
		}
		Else If CapsOption = Upper
			StringUpper, sString, sString
		Else If CapsOption = Lower
			StringLower, sString, sString
		Else If CapsOption = Title
			StringUpper, sString, sString, T
		
		Send {Raw}%sString%

		CSLength := StrLen(sString)
		Send {Left %CSLength%}+{Right %CSLength%}

	;Empty Variables
	CSLength := "", ClipboardTemp := "", CapsOption := "", sString := "", sAlphaString := ""
}

/*
PRIVATE FUNCTIONS
(Not intended to be used outside of CaseText)
*/

;PRIVATE_CaseText_TextOnly returns a string that removes all numbers, puncuation, and spaces 
;(and anything else that is not "is alpha").

PRIVATE_CaseText_TextOnly(sString)
{
	sTempText := sString
	Loop % StrLen(sTempText)
	{
		sChar := Substr(sString,A_Index,1)
		if sChar is not alpha
			StringReplace, sTempText, sTempText, %sChar%,, All
	}
	sString := "", sChar := ""
	return sTempText
}