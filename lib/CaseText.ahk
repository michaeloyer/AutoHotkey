/*
This Script is used to capitalized highlighted strings via
The Windows Clipboard.  The 4 Case options available are:

CaseTextNext()
CaseTextUpper()
CaseTextLower()
CaseTextTitle()

*/

CaseTextUpper()
{
	text := PRIVATE_CaseText_GetSelectedText()
	PRIVATE_CaseText_SendText(Format("{:U}", text))
}

CaseTextLower()
{
	text := PRIVATE_CaseText_GetSelectedText()
	PRIVATE_CaseText_SendText(Format("{:L}", text))
}

CaseTextTitle()
{
	text := PRIVATE_CaseText_GetSelectedText()
	PRIVATE_CaseText_SendText(Format("{:T}", text))	
}

CaseTextNext()
{
	text := PRIVATE_CaseText_GetSelectedText()
	alphabetText := RegExReplace(text, "[^A-Za-z]")

	if alphabetText is upper
		PRIVATE_CaseText_SendText(Format("{:L}", text))
	else if alphabetText is lower
		PRIVATE_CaseText_SendText(Format("{:T}", text))
	else
		PRIVATE_CaseText_SendText(Format("{:U}", text))
}

PRIVATE_CaseText_SendText(text)
{
	if not text 
	{
		MsgBox 48, CaseText Error, CaseText cannot be used on blank text.
		return
	}

	if RegExMatch(text, "[\n\r]")
	{
		MsgBox 48, CaseText Error, CaseText cannot be used on multi-line text.
		return
	}

	Send %text%

	textLength := StrLen(text)
	Send {Left %textLength%}+{Right %textLength%}
}

PRIVATE_CaseText_GetSelectedText()
{
	clipboardTemp := ""
	clipboardTemp := ClipboardAll
	Clipboard := ""
	
	Send ^c
	ClipWait, 1
	if Not Clipboard 
		return
		
	Clipboard := Clipboard ;Converting clipboard to plain text
	selectedText := Clipboard 
	Clipboard := clipboardTemp ;Return what was previously on the clipboard

	return selectedText 
}