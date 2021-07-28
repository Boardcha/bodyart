<%
extraChars = ""

Function getRandomNum(lbound, ubound) 
For j = 1 To (250 - ubound)
	Randomize 
	getRandomNum = Int(((ubound - lbound) * Rnd) + 1)
Next 
End Function

Function getRandomTokenChar(number, lower, upper, extra) 
numberChars = "0123456789"
lowerChars = "abcdefghijklmnopqrstuvwxyz"
upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
charSet = extra
	if (number = "true") Then charSet = charSet + numberChars
	if (lower = "true") Then charSet = charSet + lowerChars
	if (upper = "true") Then charSet = charSet + upperChars
jmi = Len(charSet) 
getRandomTokenChar = Mid(charSet, getRandomNum(1, jmi), 1)
End Function

Function getToken(length, extraChars)
rc = ""
	If (length > 0) Then
		rc = rc + getRandomTokenChar("true", "true", "true", extraChars)
	End If

	For idx = 1 To length - 1
		rc = rc + getRandomTokenChar("true", "true", "true", extraChars)
	Next

getToken = rc

End Function

' To use call function below, along with the length you want the string to be. Extra characters are only if desired -- not currently using that

'  getToken(32, extraChars)
%>