answer=""
default=""

firstString=""
lastString=""
blank=""

prevFirstString=""
prevBlank=""
prevLastString=""

x=MsgBox("Welcome to Study Helper.", 0+64, "Study Helper")

continue=6
keepStrings=7

Do While continue=6
	If keepStrings=7 Then
		If Not prevFirstString="" Then
			showStrings=MsgBox("Show input as Strings are entered?", 4+32, "Set Options:")
		End If
		
		prevFString=firstString
		
		firstString=""
		lastString=""
		blank=""
		
		If showStrings=6 Then
			default=""
		Else
			default=prevFirstString
		End If
		Do While firstString=""
			firstString=InputBox("Enter the first part of the fill-in-the-blank sentence:", "Getting Input:", default)
		Loop
		If showStrings=6 Then
			default=firstString
		Else
			default=prevBlank
		End If
		Do While blank=""
			blank=InputBox("Enter the answer WITHOUT capital letters:", "Getting Input:", default)
		Loop
		If showStrings=6 Then
			default=firstString & " " & blank
		Else
			default=prevLastString
		End If
		Do While lastString=""
			lastString=InputBox("Enter the last part of the fill-in-the-blank sentence:", "Getting Input:", default)
		Loop
	End If
	Do While (Not (answer="SKIP" Or answer=blank))
		answer=InputBox(firstString & " (BLANK) " & lastString, "Fill in the blank or type SKIP:", answer)
	Loop
	If answer=blank Then
		continue=MsgBox("Congratulations! You got the correct answer! Try again?", 4+32, "Congratulations!")
		If continue=6 Then
			answer=""
			keepStrings=MsgBox("Keep previous Strings?", 4+32, "What should I do?")
			If keepStrings=7 Then
				showStrings=7
				
				prevFirstString=firstString
				prevBlank=blank
				prevLastString=lastString
			End If
		End If
	Else
		x=MsgBox("The answer was " & blank & ".", 0+64, "Skipped")
		continue=MsgBox("Would you like to try again?", 4+32, "Try again?")
		If continue=6 Then
			answer=""
			keepStrings=MsgBox("Keep previous Strings?", 4+32, "What should I do?")
			If keepStrings=7 Then
				showStrings=7
				
				prevFirstString=firstString
				prevBlank=blank
				prevLastString=lastString
			End If
		End If
	End If
Loop
