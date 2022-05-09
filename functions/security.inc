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
%>