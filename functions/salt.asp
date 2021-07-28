<%
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
otherChars = "`~!@#$%^&*()-_=+[{]}\\|;:\,<.>/?"
charSet = extra
	if (number = "true") Then charSet = charSet + numberChars
	if (lower = "true") Then charSet = charSet + lowerChars
	if (upper = "true") Then charSet = charSet + upperChars
	if (other = "true") Then charSet = charSet + otherChars
jmi = Len(charSet) 
getRandomChar = Mid(charSet, getRandomNum(1, jmi), 1)
End Function

Function getSalt(length, extraChars)
rc = ""
	If (length > 0) Then
		rc = rc + getRandomChar("true", "true", "true", "true", extraChars)
	End If

	For idx = 1 To length - 1
		rc = rc + getRandomChar("true", "true", "true", "true", extraChars)
	Next

getSalt = rc

End Function

' To use call function below, along with the length you want the string to be. Extra characters are only if desired -- not currently using that

'  getSalt(32, extraChars)
%>