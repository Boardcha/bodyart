<%

	firstNumber = "true"
	firstLower =  "true"
	firstUpper =  "true"
	firstOther =  "false"
	latterNumber = "true"
	latterLower = "true"
	latterUpper = "true"
	latterOther = "false"
	passwordLength = 25
	extraChars = ""

	Function getRandomNum(lbound, ubound)
		For j = 1 To (250 - ubound)
			Randomize
			getRandomNum = Int(((ubound - lbound) * Rnd) + 1)
		Next
	End Function

	Function getRandomChar(number, lower, upper, other, extra) 
		numberChars = "0123456789"
		lowerChars = "abcdefghijklmnopqrstuvwxyz"
		upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		otherChars = "`~!@#$%^&*()-_=+[{]}\\|;:"""'\,<.>/? "
		charSet = extra
			if (number = "true") Then charSet = charSet + numberChars
			if (lower = "true") Then charSet = charSet + lowerChars
			if (upper = "true") Then charSet = charSet + upperChars
			if (other = "true") Then charSet = charSet + otherChars
		jmi = Len(charSet) 
		getRandomChar = Mid(charSet, getRandomNum(1, jmi), 1)
	End Function

	Function getPassword(length, extraChars, firstNumber, firstLower, _
		firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)
		rc = ""
			If (length > 0) Then
				rc = rc + getRandomChar(firstNumber, firstLower, firstUpper, firstOther, extraChars)
			End If

			For idx = 1 To length - 1
				rc = rc + getRandomChar(latterNumber, latterLower, latterUpper, latterOther, extraChars)
			Next
		getPassword = rc
	End Function
	
	strRandomCode = getPassword(passwordLength, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)

%>