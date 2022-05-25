<%
	Function Sanitize(str)
		Dim remove, slice
		'String slices to remove
		'remove = Array("<script", "declare @q", "http", "src=", "CC:", "TO:", "href", "url=", ".exe", "%", "(", ")", "<", ">", """", "'")
		'Trim the string to a specified number of characters
		'If cut > 0 Then str = Left(str, cut)
		'Cycles through the list of string slices to remove
		'For Each slice In remove
		'	str = Replace(str, slice, "", 1, -1, 1)
		'Next
		str = Replace(str, "<", "&lt;")
		str = Replace(str, ">", "&gt;") 
		str = Replace(str, "%3C", "&lt;")
		str = Replace(str, "%3E", "&gt;") 		
		'Sets return function value			
		sanitize = str
	End Function
	
	Function FixCookies()
		if request.cookies("gaugecard") = "undefined" then
			response.cookies("gaugecard") = ""
		end if
		if request.cookies("freegift1id") = "undefined" then
			response.cookies("freegift1id") = ""
		end if
		if request.cookies("freegift2id") = "undefined" then
			response.cookies("freegift2id") = ""
		end if
		if request.cookies("freegift3id") = "undefined" then
			response.cookies("freegift3id") = ""
		end if
		if request.cookies("freegift4id") = "undefined" then
			response.cookies("freegift4id") = ""
		end if
		if request.cookies("freegift5id") = "undefined" then
			response.cookies("freegift5id") = ""
		end if
		if request.cookies("freegift6id") = "undefined" then
			response.cookies("freegift6id") = ""
		end if
		if request.cookies("freegift7id") = "undefined" then
			response.cookies("freegift7id") = ""
		end if
	End Function
		
	Call FixCookies()	
%>