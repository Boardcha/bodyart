<%
extraChars = ""

Function getRandomNum(lbound, ubound) 
For j = 1 To (250 - ubound)
	Randomize 
	getRandomNum = Int(((ubound - lbound) * Rnd) + 1)
Next 
End Function

Function getRandomTokenChar(number) 
numberChars = "0123456789"
charSet = extra
	if (number = "true") Then charSet = charSet + numberChars
jmi = Len(charSet) 
getRandomTokenChar = Mid(charSet, getRandomNum(1, jmi), 1)
End Function

Function getInteger(length)
rc = ""
	If (length > 0) Then
		rc = rc + getRandomTokenChar("true")
	End If

	For idx = 1 To length - 1
		rc = rc + getRandomTokenChar("true")
	Next

getInteger = rc

End Function

' To use call function below, along with the length you want the string to be. Extra characters are only if desired -- not currently using that

'  getInteger(10, extraChars)
%>