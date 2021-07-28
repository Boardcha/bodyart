<%
'
' ADOBE SYSTEMS INCORPORATED
' Copyright 2007 Adobe Systems Incorporated
' All Rights Reserved
' 
' NOTICE:  Adobe permits you to use, modify, and distribute this file in accordance with the 
' terms of the Adobe license agreement accompanying it. If you have received this file from a 
' source other than Adobe, then your use, modification, or distribution of it requires the prior 
' written permission of Adobe.
'

'
'	Copyright (c) InterAKT Online 2000-2005
'
	' C-style sprintf function
	' Returns a string with %s substituted with values from the array
	Function KT_sprintf (strSource, ByRef arrParams)
			Dim strRest: strRest = strSource
			Dim strReturn: strReturn = ""
			
			For i = 0 To UBound(arrParams)
				pos = Instr(strRest, "%s")
				If pos <> 0 Then
					strReturnSlice = Left(strRest, pos+1)
					strReturnSlice = Replace(strReturnSlice, "%s", arrParams(i))
					strReturn = strReturn & strReturnSlice					
					strRest = Mid(strRest, pos+2)
				End If
			Next
			strReturn = strReturn & strRest
			KT_sprintf = strReturn
	End Function



	Function KT_isSet (ByRef var)
		KT_isSet = True
		If IsObject(var) Then
			If var is Nothing Then
				KT_isSet = False
				Exit Function
			ElseIf IsEmpty(var) Then
				KT_isSet = False
				Exit function
			End If	
		ElseIf IsArray(var) Then
			KT_isSet = True
			Exit function
		Else
			If IsNull(var) Then 
				KT_isSet = False
				Exit function
			ElseIf IsEmpty(var) Then
				KT_isSet = False
				Exit function
			End If
		End If
	End Function
	
	
	Function KT_cloneObject (ByRef objToClone)
		Set objToReturn = Server.CreateObject("Scripting.Dictionary")		
		For each it in objToClone
			If isObject (objToClone(it)) Then
			   Set objToReturn(it) = KT_cloneObject(objToClone(it))
			Else
				objToReturn(it) = objToClone(it)
			End If   
		Next
		Set KT_cloneObject = objToReturn
	End Function

	
	' Checks if the array contains a specified value
	' The comparation can be case sensitive or insensitive
	Function KT_in_array (mSearchValue, ByRef mArray, caseSensitive)
		compareType = 1
		If caseSensitive Then
			compareType = 0
		End If
		Dim mReturn: mReturn = False
		If IsArray(mArray) Then
			Dim i
			For i = LBound(mArray) To UBound(mArray)
				If StrComp(mSearchValue, mArray(i), compareType) = 0 Then
					mReturn = True
					Exit for
				End If				
			Next
		End If		
		KT_in_array = mReturn
	End Function

	' Checks if the array contains a specified value
	' The comparation can be case sensitive or insensitive
	Function KT_indexOf (mSearchValue, ByRef mArray, caseSensitive)
		KT_indexOf = -1
		compareType = 1
		If caseSensitive Then
			compareType = 0
		End If
		Dim mReturn: mReturn = False
		If IsArray(mArray) Then
			Dim i
			For i = LBound(mArray) To UBound(mArray)
				If StrComp(mSearchValue, mArray(i), compareType) = 0 Then
					KT_indexOf = i
				End If				
			Next
		End If		
	End Function


	Function KT_array_keys(ByRef collection)
		KT_array_keys = collection.Keys
	End Function

	' Adds a new item to an array 
	' Returns the modified array
	Function KT_array_push (ByRef m_array, m_item )
		Dim initialsize : initialsize = ubound(m_array)
		ReDim preserve m_array(initialsize+1)
		m_array(initialsize+1) = m_item
		KT_array_push = m_array
	End Function

	
	Function KT_preg_match(byval m_Pattern, byval m_Subject, byref m_Matches)
		ExecuteGlobal "" & _
		"If Not isObject(regEx__global) Then" & vbNewLine & _
		"	Set regEx__global = New RegExp" & vbNewLine & _
		"	regEx__global.Global = True"  & vbNewLine & _
		"	regEx__global.MultiLine = True"  & vbNewLine & _
		"	regEx__global.IgnoreCase = True"  & vbNewLine & _
		"End If"	
		regEx__global.Pattern = m_Pattern
				
		Set m_Matches = regEx__global.Execute(m_Subject)
		Set regEx = Nothing
	End Function
	
	
	Function KT_preg_replace(byval m_Pattern, byval m_Replacement, byval m_Subject)
	
		ExecuteGlobal "" & _
		"If Not isObject(regEx__global) Then" & vbNewLine & _
		"	Set regEx__global = New RegExp" & vbNewLine & _
		"	regEx__global.Global = True"  & vbNewLine & _
		"	regEx__global.MultiLine = True"  & vbNewLine & _
		"	regEx__global.IgnoreCase = True"  & vbNewLine & _
		"End If"	
		regEx__global.Pattern = m_Pattern
   
		KT_preg_replace = regEx__global.Replace(m_Subject, m_Replacement)   
		Set regEx = Nothing
	End Function	
	
	
	Function KT_preg_test(byval m_Pattern, byval m_Subject)
		ExecuteGlobal "" & _
		"If Not isObject(regEx__global) Then" & vbNewLine & _
		"	Set regEx__global = New RegExp" & vbNewLine & _
		"	regEx__global.Global = True"  & vbNewLine & _
		"	regEx__global.MultiLine = True"  & vbNewLine & _
		"	regEx__global.IgnoreCase = True"  & vbNewLine & _
		"End If"
		regEx__global.Pattern = m_Pattern
		KT_preg_test = regEx__global.Test(m_Subject)
	End Function
	
	Function KT_preg_split(m_charsDelims, m_Subject, m_Limit, m_Flag)
		str = m_Subject
		If isNull(m_Flag) Then
			m_Flag = ""
		End If
		
		pos = instr(m_charsDelims, "/")
		If pos <> 0 Then
			placeholderChar = mid(m_charsDelims, pos, 1)
		Else
			placeholderChar = "`"	
		End If
		str = KT_preg_replace(m_charsDelims, placeholderChar, str)
		
		Dim retSplit
		retSplit = split(str, placeholderChar)
		If m_Flag = "PREG_SPLIT_NO_EMPTY" Then
			Dim ret, i
			ret = array()
			
			For i = 0 to ubound(retSplit)
				If retSplit(i) <> "" Then
					ret = KT_array_push (ret, retSplit(i))
				End If
			Next
			KT_preg_split = ret
			Exit Function
		End If
		KT_preg_split = retSplit
	End Function


	Function KT_preg_quote(m_Subject)
		KT_preg_quote = KT_preg_replace("([\.\\\+\*\?\[\^\]\$\(\)\{\}\=\!\<\>\|\:])", "\$1", m_Subject)
	End Function
	
	
	Function KT_strip_tags_with_allowables (ByRef stripText, allowable_tags)
		Dim text: text = stripText
		allowable_tags = replace (allowable_tags, "><", ",")
		allowable_tags = replace (allowable_tags, "<", "")
		allowable_tags = replace (allowable_tags, ">", "")
		
		arr_allowable_tags = split(allowable_tags, ",")

		Dim matches
		KT_preg_match "<(/?)(\w+)[^>]*>", text, matches
		For each match in matches
			t = match.Value
			tag_name = KT_preg_replace ("<(/?)(\w+)[^>]*>", "$2", t)
			tag_name = lcase(tag_name)
			If Not KT_in_array(tag_name, arr_allowable_tags, false) Then
				text = replace(text, t, "")
			End If
		Next
		KT_strip_tags_with_allowables = text
	End Function

	Function KT_strip_tags (ByRef stripText)
		KT_strip_tags = stripText
		KT_strip_tags = KT_strip_tags_with_allowables(KT_strip_tags, "")
	End Function


	' retrieves the url path to the root site (eg /sites/MYSITE )
	Function KT_GetSitePath()
		If isEmpty(Session("KT_SitePath")) Then
			KT_SetPathSessions
		End If
		KT_GetSitePath = Session("KT_SitePath")
	End Function


	Function KT_GetURLToResource(urlRelativeToRootSite)
		If Session("KT_URLToRootSite") <> "" Then
			KT_GetURLToResource = Session("KT_URLToRootSite") & "/" & urlRelativeToRootSite
			Exit Function
		End If
		
		If isEmpty(Session("KT_SitePath")) Then
			KT_SetPathSessions
		End If
		
		url = KT_getServerName() & Session("KT_SitePath") 
		Session("KT_URLToRootSite") = url
		
		KT_GetURLToResource = url & "/" & urlRelativeToRootSite
	End Function


Function KT_escapeAttribute(val)
        If not KT_isSet(val) Then
            KT_escapeAttribute = "" 
            Exit Function
        End If
        On Error Resume Next
        If TypeName(val) = "Boolean" Then
            If err.Number = 0 Then
                If val Then
                    KT_escapeAttribute = "1"
                Else
                    KT_escapeAttribute = "0" 
                End If
                On Error GoTo 0
                Exit Function
            End If 
        End If
        On Error GoTo 0
        KT_escapeAttribute = val
        KT_escapeAttribute = replace(KT_escapeAttribute, """", "&quot;")
        KT_escapeAttribute = replace(KT_escapeAttribute, "<", "&lt;")
        KT_escapeAttribute = replace(KT_escapeAttribute, ">", "&gt;")
    End Function 

	Function KT_FormatForList(value, maxChars)
		myVal = value & ""
		If Trim(myVal) = "" Then
			KT_FormatForList = "&nbsp;"
			Exit Function	
		End If
		isBigger = false
		myVal = KT_preg_replace("<head[^>]*>[\w\W]*?<\/head>[" & vbNewLine & "]*", "" , myVal)
		myVal = KT_preg_replace("<link[^>]*>[" & vbNewLine & "]*", "" , myVal)
		myVal = KT_preg_replace("<script[^>]*>[\w\W]*?<\/script>[" & vbNewLine & "]*", "" , myVal)
		myVal = KT_preg_replace("<style[^>]*>[\w\W]*?<\/style>[" & vbNewLine & "]*", "" , myVal)
		myVal = KT_strip_tags(myVal)
								
		If maxChars <> -1 Then
			If len(trim(myVal)) > maxChars Then
				myVal = left(myVal, maxChars)
				isBigger = true
			End If
		End If
		myVal = replace(replace(myVal, "<", "&lt;"), ">", "&gt;")
		myVal = trim(myVal)
		If myVal = "" Then
			myVal = "&nbsp;"
		End If
		If isBigger Then
			myVal = myVal & "..."
		End If
		KT_FormatForList = myVal
	End Function
	
	Function KT_escapeJS(val)
		KT_escapeJS = KT_addcslashes(val, "\'")
		KT_escapeJS = replace(KT_escapeJS, vbNewLine, "\n")
		KT_escapeJS = replace(KT_escapeJS, vbCr, "\n")
		KT_escapeJS = replace(KT_escapeJS, vbTab, "\t")
	End Function
	
	Function KT_addReplaceParam(qstring__param, paramName__param, paramValue__param)
		Dim qstring, paramName, paramValue
		qstring = qstring__param
		paramName = paramName__param
		paramValue = paramValue__param
		
		' extract the URI if any
		If instr(qstring, "?") <> 0 Then
			uri = KT_preg_replace("\?.*$", "?", qstring)
			qstring = KT_preg_replace("^.*\?", "", qstring)
		Else
			If instr(qstring, "=") <> 0 Then
				uri = ""
			Else
				uri = qstring
				If paramValue <> ""  Then
					uri = uri & "?"
				End If
				qstring = ""
			End If
		End If
		
		' the list of parameters
		arr = split(qstring, "&")
		Dim i
		Dim found: found = false
		newArr = array()
		For i=0 to Ubound(arr) 
			If arr(i) <> "" Then
				tmpArr = split(arr(i), "=")
				If tmpArr(0) = paramName Then
					' replace paramName
					If paramValue <> "" Then
						param = paramName & "=" & Server.URLEncode(paramValue)
						newArr = KT_array_push(newArr, param)
					End If	
					found = true
				ElseIf left(paramName, 1) = "#" Then
					' remove the param using regular expressions
					If KT_preg_test(mid(paramName,2),tmpArr(0)) and paramValue = "" Then
						found = true							
					Else
						' leave it unmodified
						val = ""
						If ubound(tmpArr) = 1 Then
							val = tmpArr(1)
						End If	
						param = tmpArr(0) & "=" & val	
						newArr = KT_array_push(newArr, param)
					End If	
				Else 
					val = ""
					If ubound(tmpArr) = 1 Then
						val = tmpArr(1)
					End If	
					param = tmpArr(0) & "=" & val
					newArr = KT_array_push(newArr, param)
				End If
			End If	
		Next
		
		' add paramName to the list if not replaced
		If not found Then
			If left(paramName, 1) <> "#" And paramValue <> "" Then
				param =  paramName & "=" & Server.URLEncode(paramValue)
				newArr = KT_array_push(newArr, param)
			End If	
		End If
		
		ret = join (newArr, "&")
		ret = KT_preg_replace("^&", "", ret)
		
		' if no parameters, remove the trailing ? from the URI
		If ret = "" Then
			uri = KT_preg_replace("\?$", "", uri)	
		End If
		' merge the URI with the new list
		ret = uri & ret 
		
		KT_addReplaceParam = ret
	End Function
	
	
	Function KT_redir(url)
		Response.Redirect(url)
		Response.End()
	End Function
	
 '
 ' Converts a date/time/datetime from screen format to database format
 ' For internal use
 ' @param string $date The date in screen format
 ' returns string the date in database format
 '	
	Function KT_formatDate2DB(date__param)
		If Not KT_isSet(date__param) Then
			KT_formatDate2DB = ""
			Exit Function
		End If
		KT_formatDate2DB = KT_convertDateCall(date__param, null)
	End Function

'
' Converts a date/time/datetime from database format to screen format
' Used for date display
' @param string $date The date in database format
' returns string the date in screen format
'
	Function KT_formatDate(date__param) 
		If Not KT_isSet(date__param) Then
			KT_formatDate = ""
			Exit Function
		End If	
		KT_formatDate = KT_convertDateCall(date__param, "toscreen")
	End Function

	Sub KT_getInternalTimeFormat()
		KT_db_date_format = replace(KT_db_date_format, "Y", "y", 1, -1, 0)
		KT_db_date_format = replace(KT_db_date_format, "M", "m", 1, -1, 0)
		KT_db_date_format = replace(KT_db_date_format, "D", "d", 1, -1, 0)
		KT_screen_date_format = replace(KT_screen_date_format, "Y", "y", 1, -1, 0)
		KT_screen_date_format = replace(KT_screen_date_format, "M", "m", 1, -1, 0)
		KT_screen_date_format = replace(KT_screen_date_format, "D", "d", 1, -1, 0)
		ExecuteGlobal "KT_screen_time_format_internal = replace(KT_screen_time_format, ""m"", ""i"")"
		ExecuteGlobal "KT_db_time_format_internal = replace(KT_db_time_format, ""m"", ""i"")"
	End Sub
	
	Function KT_convertDateCall(date__param, toScreen)
		If Not KT_preg_test("^([\d-\/\[\]\(\)\s\*\|\+\.:=,]|a|p|am|pm)+$", date__param) Then
			KT_convertDateCall = date__param
			Exit Function
		End If	
		If isNull(toScreen) Then
			toScreen = ""
		End If
		KT_getInternalTimeFormat
		If KT_preg_test("^\d+$", date__param) Then
			If len(trim(date__param)) = 14 Then
				newDate = mid(date__param, 1, 4) & "-" & mid(date__param, 5,2) & "-" & mid(date__param, 7,2) 
				newDate = newDate & " " & mid(date__param, 9,2) & ":" & mid(date__param, 11,2) & ":" & mid(date__param, 13,2)
				date__param = newDate			
				fromFormat = "yyyy-mm-dd HH:ii:ss"
				toFormat = KT_screen_date_format &  " " & KT_screen_time_format_internal
			Else
				KT_convertDateCall = date__param
				Exit Function
			End If	
		End If

		If (Not KT_preg_test("^(\d+[-\/\[\]\(\)\s\*\|\+\.=,]\d+[-\/\[\]\(\)\s\*\|\+\.=,]\d+)+", date__param)) And (Not KT_preg_test("^\d+:\d+(:\d+|\s+a|\s+p|\s+am|\s+pm)", date__param)) Then
			KT_convertDateCall = date__param
			Exit Function
		End If

		If toScreen = "" And KT_preg_test("^\d+[-\/\[\]\(\)\s\*\|\+\.=,]\d+[-\/\[\]\(\)\s\*\|\+\.=,]\d+\s+\d+", date__param) Then
			date__param = KT_expandTime(date__param)
		End If
		
		
		If instr(date__param, " ") <> 0 And instr(date__param, ":") <> 0 and instr(date__param, " ") < instr(date__param, ":") Then
			fromFormat = KT_screen_date_format & " " & KT_screen_time_format_internal
			toFormat = KT_db_date_format &  " " & KT_db_time_format_internal
		ElseIf instr(date__param, ":") <> 0 Then
			fromFormat = KT_screen_time_format_internal
			toFormat = KT_db_time_format_internal
		Else	
			fromFormat = KT_screen_date_format
			toFormat = KT_db_date_format
		End If
		
		If toScreen = "toscreen" Then
			tmp = fromFormat
			fromFormat = toFormat
			toFormat = tmp
			
			' HACK
			' because ADODB does some extra conversion from db format to a LOCALE format.. must overcome this problem
			' rs(datefield) is not in db format. it's in locale format
			' must do a conversion from this LOCALE to db format .. before doing the normal conversion db format -> screen format
			
			On Error Resume Next
			ADO_db_date = Cdate(date__param)
			' getting date
			Dim yyyy: yyyy = CInt(DatePart("yyyy", ADO_db_date))
			Dim m: m = CInt(DatePart("m",  ADO_db_date))
			If m < 10 Then
				m = "0" & m
			End If
			Dim d: d = CInt(DatePart("d", ADO_db_date))
			If d < 10 Then
				d = "0" & d
			End If
			
			' getting time
			Dim h: h = Cint(Hour(ADO_db_date))
			If h < 10 Then
				h = "0" & h
			End If
			Dim min: min = Cint(Minute(ADO_db_date))
			If min < 10 Then
				min = "0" & min
			End If
			Dim sec: sec = Cint(Second(ADO_db_date))
			If sec < 10 Then
				sec = "0" & sec
			End If	

			db_date = yyyy & "-" & m & "-" & d & " " & h & ":" & min & ":" & sec
			If Cstr(yyyy) = "1900" And m = "01" And d = "01" Then
				toFormat = KT_screen_time_format_internal
			End If
			If Err.Number = 0 Then
				date__param = KT_convertDate(db_date, "yyyy-mm-dd HH:ii:ss", toFormat)
			End If
			On Error GoTo 0
			
			' check typed format
			checkTypeFormat = KT_VBRegexp(KT_date2regexp(toFormat))
			If not KT_preg_test(checkTypeFormat, date__param) Then
				KT_convertDateCall = date__param
			Else
				KT_convertDateCall = KT_stripTime(date__param)	
			End If
			Exit Function
		End If
		
		KT_convertDateCall = KT_convertDate(date__param, fromFormat, toFormat)
	End Function


	
	 ' Strips empty values from time expressions
	 ' @param string $date - datetime expression
	 ' return new datetime without 0's
	 '
	Function KT_stripTime(date__param)
		If instr(date__param, " ") <> 0 And instr(date__param, ":") <> 0 and instr(date__param, " ") < instr(date__param, ":") Then
			pos_space = instr(date__param, " ")
			time_value = trim(mid(date__param, pos_space+1))
			date_value = left(date__param, pos_space-1)			
			
			timeArr = array()
			If KT_preg_test("^(\d+)(:)?(\d*)(:)?(\d*)(:)?\s*(am|pm|a|p)?$", time_value) Then
				formated_time_value = KT_preg_replace("^(\d+)(:)?(\d*)(:)?(\d*)(:)?\s*(am|pm|a|p)?$", "$1:$3:$5# $7", time_value)
				time_part = left(formated_time_value, instr(formated_time_value, "#")-1)
				ampm_part = mid(formated_time_value, instr(formated_time_value, "#")+1)
				
				timeArr = split(time_part, ":")
			End If				
			length = ubound(timeArr)
			For i = length to 0 step -1
				If timeArr(i) <> "0" And timeArr(i) <> "00" And timeArr(i) <> "" Then
					Exit For
				Else	
					If ubound(timeArr) <> 0 Then
						Redim Preserve timeArr(ubound(timeArr)-1)
					Else
						timeArr = array()
					End If
				End If
			Next
			time_value = join(timeArr, ":")
			If time_value <> "" Then
				KT_stripTime = date_value & " " & time_value & rtrim(ampm_part)
			Else
				KT_stripTime = date_value
			End If
			Exit Function
		End If
		KT_stripTime = date__param
	End Function
	
	'
	' Expands time expressions to full screen format
	' @param string $date - datetime expression
	' return new datetime with full time part
	'
	Function KT_expandTime(date__param)
		pos_space = instr(date__param, " ")
		time_value = trim(mid(date__param, pos_space+1))
		date_value = left(date__param, pos_space-1)			
		
		str_hour = "00"
		str_min = "00"
		str_sec = "00"
		If KT_preg_test("^(\d+)(:)?(\d*)(:)?(\d*)(:)?\s*(am|pm|a|p)?$", time_value) Then
			formated_time_value = KT_preg_replace("^(\d+)(:)?(\d*)(:)?(\d*)(:)?\s*(am|pm|a|p)?$", "$1:$3:$5#$7", time_value)
			time_part = left(formated_time_value, instr(formated_time_value, "#")-1)
			ampm_part = mid(formated_time_value, instr(formated_time_value, "#")+1)
			
			timeArr = split(time_part, ":")
			If timeArr(0) <> "" Then
				str_hour = timeArr(0)
			End If

			If timeArr(1) <> "" Then
				str_min = timeArr(1)
			End If

			If timeArr(2) <> "" Then
				str_sec = timeArr(2)
			End If
		End If				
		
		If (lcase(ampm_part) = "p" Or lcase(ampm_part) = "pm") And Cint(str_hour) < 12 Then
			str_hour = Cstr(12 + Cint(str_hour))
		End If
		
		If (lcase(ampm_part) = "a" Or lcase(ampm_part) = "am") And Cint(str_hour) = 12 Then
			str_hour = "00"
		End If

	
		KT_expandTime = date_value & " " & KT_convertDate(str_hour  & ":"  & str_min & ":" & str_sec, "HH:ii:ss", KT_screen_time_format_internal)
	End Function
	
	
	Function KT_convertDate(date__param, inFmt, outFmt)
		If inFmt = "" Or outFmt = "" Or inFmt = outFmt Then
			KT_convertDate = date__param
		End If
		If date__param = "" Then
			KT_convertDate = ""
			Exit Function
		End If
			
		If not isObject(KT_convertDate_GLOBAL_CACHE) Then
			ExecuteGlobal "Set KT_convertDate_GLOBAL_CACHE = Server.CreateObject(""Scripting.Dictionary"")" & vbNewLine & _
					"Set KT_convertDate_GLOBAL_CACHE(""inout_format"") = Server.CreateObject(""Scripting.Dictionary"")" & vbNewLine & _
					"Set KT_convertDate_GLOBAL_CACHE(""output"") = Server.CreateObject(""Scripting.Dictionary"")" & vbNewLine
		End If
		ExecuteGlobal "Set CACHER = KT_convertDate_GLOBAL_CACHE"
		
		If CACHER("inout_format").Exists(inFmt) Then
			Set inFmtRule = CACHER("inout_format")(inFmt)
		Else
			Set inFmtRule = KT_format2rule(inFmt)
			Set CACHER("inout_format")(inFmt) = inFmtRule
		End If
		
		If CACHER("inout_format").Exists(outFmt) Then
			Set outFmtRule = CACHER("inout_format")(outFmt)
		Else
			Set outFmtRule = KT_format2rule(outFmt)
			Set CACHER("inout_format")(outFmt) = outFmtRule
		End If

		If CACHER("output").Exists(outFmt) Then
			outRule = CACHER("output")(outFmt)
		Else
			outRule = KT_format2outRule(outFmt)
			CACHER("output")(outFmt) = outRule
		End If
		
		Set dateArr = KT_applyDate2rule(date__param, inFmtRule)
		outdate = KT_applyOutRule2date(dateArr, outFmtRule, outRule)
		KT_convertDate = outdate
	End Function


'
' Removes extra chars from a date format, in order to obtain a parsable definition
' @param string $format The format to be stripped
'
	Function KT_format2outRule(format__param)
		format = format__param
		format = replace(format, "yyyy", "y")
		format = replace(format, "yy", "y")
		format = replace(format, "mm", "m")
		format = replace(format, "dd", "d")						
		
		format = replace(format, "hh", "h")
		format = replace(format, "HH", "H")
		format = replace(format, "ii", "i")
		format = replace(format, "ss", "s")						
		format = replace(format, "tt", "t")								
		KT_format2outRule = format
	End Function

'
' Splits a date format into a chunked representation
' @param string $format The format to be precessed
' returns array the format in a chunked form (with chunks position and length)
'
	Function KT_format2rule (format)
		Dim rule: Set rule = Server.CreateObject("Scripting.Dictionary")
		' simulate 
		rulePieces = KT_preg_split("[-\/\[\]\(\)\s\*\|\+\.:=,]", format, -1, null)
		Dim i
		For i=0 to ubound(rulePieces)
			rulePiece = rulePieces(i)
			If rulePiece = "yyyy" Or rulePiece = "yy" Or rulePiece = "y" Then
				Set rule("y") = Server.CreateObject("Scripting.Dictionary")
				rule("y")("piece") = i
				rule("y")("len") = len(rulePiece)
			End If
			If rulePiece = "mm" Or rulePiece = "m" Then
				Set rule("m") = Server.CreateObject("Scripting.Dictionary")
				rule("m")("piece") = i
				rule("m")("len") = len(rulePiece)
			End If			
			If rulePiece = "dd" Or rulePiece = "d" Then
				Set rule("d") = Server.CreateObject("Scripting.Dictionary")
				rule("d")("piece") = i
				rule("d")("len") = len(rulePiece)
			End If
			If rulePiece = "HH" Or rulePiece = "H" Then
				Set rule("H") = Server.CreateObject("Scripting.Dictionary")
				rule("H")("piece") = i
				rule("H")("len") = len(rulePiece)
			End If
			If rulePiece = "hh" Or rulePiece = "h" Then
				Set rule("h") = Server.CreateObject("Scripting.Dictionary")
				rule("h")("piece") = i
				rule("h")("len") = len(rulePiece)
			End If	
			If rulePiece = "ii" Or rulePiece = "i" Then
				Set rule("i") = Server.CreateObject("Scripting.Dictionary")
				rule("i")("piece") = i
				rule("i")("len") = len(rulePiece)
			End If						
			If rulePiece = "ss" Or rulePiece = "s" Then
				Set rule("s") = Server.CreateObject("Scripting.Dictionary")
				rule("s")("piece") = i
				rule("s")("len") = len(rulePiece)
			End If
			If rulePiece = "tt" Or rulePiece = "t" Then
				Set rule("t") = Server.CreateObject("Scripting.Dictionary")
				rule("t")("piece") = i
				rule("t")("len") = len(rulePiece)
			End If											
		Next
		Set KT_format2rule = rule
	End Function

'
' Splits a date into a chunked representation
' @param string $date The date to be precessed
' @param array $rule Associative array containing the date chunks order ('y'=> 1, 'm' => 3, etc)
' returns array the date in a chunked form, containig yyyy, mm, dd, HH, ii and ss
'
	Function KT_applyDate2rule(date__param, ByRef rule)
		Dim dateArr: Set dateArr = Server.CreateObject("Scripting.Dictionary")
		dateArr("y") = ""
		dateArr("m") = ""
		dateArr("d") = ""
		dateArr("H") = "00"
		dateArr("i") = "00"
		dateArr("s") = "00"
		
		datePieces = KT_preg_split("[-\/\[\]\(\)\s\*\|\+\.:=,]", date__param, -1, "PREG_SPLIT_NO_EMPTY")
		If ubound(datePieces) <> -1 Then
			For each ruleKey in rule
				Set ruleValue = rule(ruleKey)
				index = ruleValue("piece")
				If index <= ubound(datePieces) Then
					dateArr(ruleKey) = datePieces(index)
				End If
			Next
		End If
				
		ruleKeys = KT_array_keys(rule)
		If KT_in_array("h", ruleKeys, true) AND dateArr.Exists("h") Then
			dateArr("H") = dateArr("h")
			dateArr.remove "h"
		End If
		If KT_in_array("t", ruleKeys, true) Then
			If KT_isSet(dateArr("t")) Then
				value = dateArr("t")
			Else
				value = "A"	
			End If
			If UCase(left(value,1)) = "P" Then
				If Cint(dateArr("H")) < 12 Then
					dateArr("H")  = Cint(DateArr("H")) + 12
				End If
			Else
				If Cint(dateArr("H")) = 12 Then
					dateArr("H") = 0
				End If	
			End If
			dateArr.remove "t"
		End If
		
		pieces = array("y", "m", "d", "H", "i", "s")
		Dim i
		For i = 0 to Ubound(pieces) 
			piece = pieces(i)
			If len(dateArr(piece)) = 1 Then
				dateArr(piece) = "0" & dateArr(piece)
			End IF
		Next
		
		If len(dateArr("y"))  = 2 Then
			If dateArr("y") < 70 Then
				dateArr("y") = "20" & dateArr("y")
			Else
				dateArr("y") = "19" & dateArr("y")
			End If
		End If	
		Set KT_applyDate2rule = dateArr
	End Function


	Function KT_array_diff(inArr, compArr)
		outArr = array()
		Dim i
		For i=0 to ubound(inArr)
			If Not KT_in_array(inArr(i), compArr, true) Then
				outArr = KT_array_push(outArr, inArr(i))
			End If
		Next
		KT_array_diff = outArr
	End Function

'
' Processes a date array in a usable format
' @param array $dateArr Associative array containing date chunks ('y'=>'2004', 'm'=>'5', etc.)
' @param array $formatRule Associative array containing the output date formatting rules ('y'=> 2 chars, 'm' => 1 char, etc)
' @param string $outStringRule Defines the output date format
' returns string the date in the $outStringRule format
'
	Function KT_applyOutRule2date(ByRef dateArr, ByRef formatRule, ByRef outStringRule)
		dateArrKeys = KT_array_keys(dateArr)
		formatRuleKeys = KT_array_keys(formatRule)

		
		preparedKeys = KT_array_diff(formatRuleKeys, dateArrKeys)
		If ubound(preparedKeys) <> -1 Then
			If KT_in_array("h", preparedKeys, true) Then
				value = dateArr("H")
				If not dateArr.Exists("h") Then
					dateArr.Add "h", value	
				End If

				If value = 0 Then
					dateArr("h") = 12
				End If
				dateArr("t") = "AM"
				If value > 12 and value < 24 Then
					dateArr("h") = value - 12
					dateArr("t") = "PM"
				End If
			End If
			If KT_in_array("t", preparedKeys, true) Then
				value = dateArr("H")
				dateArr("t") = "AM"
				If value > 11 Then
					dateArr("t") = "PM"
				End If
			End If
		End If

	
		formatRuleKeys = KT_array_keys(formatRule)	
		Dim i
		For i = 0 to UBound(formatRuleKeys)
			key = formatRuleKeys(i)
			length = formatRule(key)("len")
			value = dateArr(key)
			
			' convert from less digits to more
			' only for HH
			If len(value) < length Then
				If key = "H" Then
					dateArr(key) = "0" & value
				End If
			End If

			' convert from more digits to less
			If len(value) > length Then
				If KT_in_array(key, array("m", "d", "i", "h", "H", "s"), true) Then
					If left (value,1) = "0" Then
						dateArr(key) = mid(value, 2)
					End If
				End If
				
				If key = "y" Then
					If len(value) = 4 Then
						value = mid(value, 3)
					End If
					If length = 1 and left(value, 1) = "0" Then
						value = mid(value,2)
					End If
					dateArr(key) = value
				End If
				
				If key = "t" Then
					dateArr(key) = left(value, 1)
				End If
			End If
		Next

		myDate = outStringRule
		For each key in dateArr
			myDate = replace(myDate, key, dateArr(key), 1, -1, 0)
		Next
		myDate = trim(KT_preg_replace("[-\/\[\]\(\)\s\*\|\+\.:=,]{2,}", "", myDate))
		
		KT_applyOutRule2date = myDate
	End Function


'
' Validates a date array
' @param $dateArr the date array
' return boolean the date is valid or not
'
	Function KT_isValidDate(ByRef dateArr, fullDateTime)
		If isnull(fullDateTime) Then
			fullDateTime = true
		End If
		Dim y, m, d, H, i, s
		
		If fullDateTime Then
			If dateArr.Exists("y") Then
				On Error Resume Next
				y = Cint(dateArr("y")) 
				If err.Number <> 0 Then
					KT_isValidDate = false
					On Error GoTo 0
					Exit Function
				End If	
				On Error GoTo 0
			Else
				KT_isValidDate = false
				Exit Function
			End If
	
			If dateArr.Exists("m") Then
				On Error Resume Next
				m  = Cint(dateArr("m")) 
				If err.Number <> 0 Then
					KT_isValidDate = false
					On Error GoTo 0
					Exit Function
				End If	
				On Error GoTo 0
				
				If m < 1 Or m >  12 Then
					KT_isValidDate = false
					Exit Function				
				End If
			Else
				KT_isValidDate = false
				Exit Function		
			End If
			
			
			maxday = KT_getDaysOfMonth(m, y)
	
			If dateArr.Exists("d") Then
				On Error Resume Next
				d = Cint(dateArr("d")) 
				If err.Number <> 0 Then
					KT_isValidDate = false
					On Error GoTo 0
					Exit Function
				End If	
				On Error GoTo 0
				
				If d < 1 Or d > maxday Then
					KT_isValidDate = false
					Exit Function		
				End If
			Else
				KT_isValidDate = false
				Exit Function				
			End If
		End If
		
		
		If dateArr.Exists("H") Then
			On Error Resume Next
			H = Cint(dateArr("H")) 
			If err.Number <> 0 Then
				KT_isValidDate = false
				On Error GoTo 0
				Exit Function
			End If	
			On Error GoTo 0
			
			If H < 0 Or H > 23 Then
				KT_isValidDate = false
				Exit Function		
			End If
		Else
			KT_isValidDate = false
			Exit Function				
		End If
	
		If dateArr.Exists("i") Then
			On Error Resume Next
			i = Cint(dateArr("i")) 
			If err.Number <> 0 Then
				KT_isValidDate = false
				On Error GoTo 0
				Exit Function
			End If	
			On Error GoTo 0
					
			If i < 0 Or i > 59 Then
				KT_isValidDate = false
				Exit Function		
			End If
		Else
			KT_isValidDate = false
			Exit Function				
		End If		

		If dateArr.Exists("s") Then
			On Error Resume Next
			s = Cint(dateArr("s")) 
			If err.Number <> 0 Then
				KT_isValidDate = false
				On Error GoTo 0
				Exit Function
			End If	
			On Error GoTo 0
					
			If s < 0 Or s > 59 Then
				KT_isValidDate = false
				Exit Function		
			End If
		Else
			KT_isValidDate = false
			Exit Function				
		End If	
		KT_isValidDate = true
	End Function


	Function KT_getDaysOfMonth (mm, yy) 
		Dim maxday
		
		If KT_in_array(Cstr(mm), Array("1", "3", "5", "7", "8", "10", "12"), False)  Then
			maxday = 31
		ElseIf KT_in_array(Cstr(mm), Array("4", "6", "9", "11"), False)  Then	
			maxday = 30
		ElseIf Cstr(mm) = "2" Then
			maxday = 28
			If ((yy mod 4 = 0) AND (yy mod 100 <> 0)) OR (yy mod 400 = 0) Then
				maxday = 29
			End If
		End If
		
		KT_getDaysOfMonth = maxday
	End Function


	Function KT_escapeExpression(expr)
		If Not isNull(expr) Then
			expr = replace(expr, """", """""")
			expr = """" & expr & """"
		Else
			expr = "null"	
		End If
		KT_escapeExpression  = expr
	End Function


	
	Function KT_addcslashes(str, charlist )
		Dim i 
		Dim outstr
		outstr = str
		For i=1 to len(charlist)
			ch2 = mid(charlist, i, 1)
			outstr = replace(outstr, ch2, "\" & ch2) 
		Next
		KT_addcslashes = outstr
	End Function
	
	Private Function KT_VBRegexp(in_regexp)
		KT_VBRegexp = KT_preg_replace ("/(.+)/[img]{0,3}", "$1", in_regexp)
	End Function	

	Private Function KT_date2regexp(in_date_format)
		txt = in_date_format
		txt = KT_preg_replace("([\/\-\.])", "DATESEPARATOR", txt)
		txt = KT_preg_quote(txt)
		txt = replace(txt, "DATESEPARATOR", "[\/\-\.]")
		txt = replace(txt, "yyyy", "([0-9]{1,4})", 1,-1,0)
		txt = replace(txt, "yy", "([0-9]{1,4})", 1,-1,0)

		txt = replace(txt, "mm", "([0-9]{1,2})", 1,-1,0)
		txt = replace(txt, "m", "([0-9]{1,2})", 1,-1,0)

		txt = replace(txt, "dd", "([0-9]{1,2})", 1,-1,0)
		txt = replace(txt, "d", "([0-9]{1,2})", 1,-1,0)
		
		txt = replace(txt, "hh", "([0-9]{1,2})", 1,-1,0)
		txt = replace(txt, "h", "([0-9]{1,2})", 1,-1,0)

		txt = replace(txt, "HH", "([0-9]{1,2})", 1,-1,0)
		txt = replace(txt, "H", "([0-9]{1,2})", 1,-1,0)

		txt = replace(txt, "ii", "([0-9]{1,2})", 1,-1,0)
		txt = replace(txt, "i", "([0-9]{1,2})", 1,-1,0)

		txt = replace(txt, "ss", "([0-9]{1,2})", 1,-1,0)
		txt = replace(txt, "s", "([0-9]{1,2})", 1,-1,0)

		txt = replace(txt, "tt", "(AM|PM|am|pm)", 1,-1,0)
		txt = replace(txt, "t", "(A|P|a|p)", 1,-1,0)
		
		txt = KT_preg_replace("(\s)", "\s", txt)
		
		txt = "/^" & txt & "$/"	
		KT_date2regexp = txt
	End Function
	
	'
	' Compares 2 date arrays
	' @param array $dateArr1
	' @param array $dateArr2
	' return integer -1, 1 or 0
	'
	Function KT_compareDates(dateArr1, dateArr2)
		time1 = dateArr1("y") & dateArr1("m") & dateArr1("d") & dateArr1("H") & dateArr1("i") & dateArr1("s")
		time2 = dateArr2("y") & dateArr2("m") & dateArr2("d") & dateArr2("H") & dateArr2("i") & dateArr2("s")
		If Cdbl(time1) > Cdbl(time2) Then
			KT_compareDates  = -1
			Exit Function
		End If
		If Cdbl(time1) < Cdbl(time2) Then
			KT_compareDates  = 1
			Exit Function
		End If
		KT_compareDates  = 0
	End Function

	Function KT_getServerName()
		serverName = ""
		port = Request.ServerVariables("SERVER_PORT")
		https = lcase(Request.ServerVariables("HTTPS"))
		If https = "on" Then
			portStr = ""
			If port <> 443 Then
				portStr = ":" & port
			End If
			serverName = "https://" & Request.ServerVariables("SERVER_NAME") & portStr
		Else
			portStr = ""
			If port <> 80 Then
				portStr = ":" & port
			End If
			serverName = "http://" & Request.ServerVariables("SERVER_NAME") & portStr
		End If
		KT_getServerName = serverName
	End Function
		
	Function KT_getUri()
		KT_getUri = KT_getServerName() & Request.ServerVariables("URL")
	End Function
	
	Function KT_getFullUri()
		Dim ret
		ret = KT_getUri()
		If Request.ServerVariables("QUERY_STRING") <> "" Then
			ret = ret & "?" &  Request.ServerVariables("QUERY_STRING")
		End If 
		KT_getFullUri = ret
	End Function

	Function KT_transformsPaths(templateUrl__param, text__param, makeAbsolut)
		templateUrl = templateUrl__param
		If templateUrl__param = "./" Then
			templateUrl = ""
		End If

		text = text__param
		scriptUri = KT_getUri()

		Dim auxArr(17)
		Dim auxArr1(16)
		Dim matches
		Dim matches2
		KT_preg_match "<(a|img|link|script|form|iframe|embed|applet)([^>]*)>", text, matches
		For each match in matches
			t = match.Value
			t2 = KT_preg_replace ("<(a|img|link|script|form|iframe|embed|applet)([^>]*)>", "$2", t)
			KT_preg_match "\s(href|src|action|codebase|archive)\s*=\s*((""((\\\""|[^""])+)"")|(\'((\\\\\'|[^\'])+)\')|([^\s]+))?", t2, matches2
			For each match2 in matches2
				tt = match2.Value
				tt2 = KT_preg_replace ("\s(href|src|action|codebase|archive)\s*=\s*((""((\\\""|[^""])+)"")|(\'((\\\\\'|[^\'])+)\')|([^\s]+))?", "$1^^^$2", tt)
				mmatch = split(tt2, "^^^")
				delim = ""
				If left(mmatch(1),1) = "'" Then
					delim = "'"
					mmatch(1) = mid(mmatch(1), 2, len(mmatch(1))-2)
				ElseIf left(mmatch(1),1) = """" Then
					delim = """"
					mmatch(1) = mid(mmatch(1), 2, len(mmatch(1))-2)
				End If 
				If  Not (lcase(mmatch(0)) = "href" And Instr(1, mmatch(1), "javascript:", 1) = 1) And Not (lcase(mmatch(0)) = "href" And Instr(mmatch(1), "#") = 1) And Not (lcase(mmatch(0)) = "action" And mmatch(1) = "")  And Not (lcase(mmatch(0)) = "href" And Instr(mmatch(1), "mailto:") = 1) And Not (lcase(mmatch(0)) = "action" And Instr(mmatch(1), "mailto:") = 1) Then
					text = KT_preg_replace(mmatch(0) & "\s*=\s*" & KT_preg_quote(delim & mmatch(1) & delim),  mmatch(0) & "=" & delim & KT_Rel2AbsUrl(scriptUri, templateUrl, mmatch(1), makeAbsolut) & delim, text)
				End If
			Next
		Next
		
		Dim matches3
		KT_preg_match "<input([^>]*)value\s*=\s*(?:(?:""(?:(?:\\\""|[^""])+)"")|(?:\'(?:(?:\\\\\'|[^\'])+)\')|(?:[^\s]+))?([^>]*)>", text, matches
		For each match in matches
			t = match.Value
			t3 = KT_preg_replace ("<input([^>]*)value\s*=\s*(?:(?:""(?:(?:\\\""|[^""])+)"")|(?:\'(?:(?:\\\\\'|[^\'])+)\')|(?:[^\s]+))?([^>]*)>", "$1^^^$2", t)
			KT_preg_match "\s(src)\s*=\s*((""((\\\""|[^""])+)"")|(\'((\\\\\'|[^\'])+)\')|([^\s]+))?", t3, matches3
			For each match3 in matches3
				tt = match3.Value
				tt3 = KT_preg_replace ("\s(src)\s*=\s*((""((\\\""|[^""])+)"")|(\'((\\\\\'|[^\'])+)\')|([^\s]+))?", "$1^^^$2", tt)
				mmatch = split(tt3, "^^^")
				delim = ""
				If left(mmatch(1),1) = "'" Then
					delim = "'"
					mmatch(1) = mid(mmatch(1), 2, len(mmatch(1))-2)
				ElseIf left(mmatch(1),1) = """" Then
					delim = """"
					mmatch(1) = mid(mmatch(1), 2, len(mmatch(1))-2)
				End If 
				text = KT_preg_replace(mmatch(0) & "\s*=\s*" & KT_preg_quote(delim & mmatch(1) & delim),  mmatch(0) & "=" & delim & KT_Rel2AbsUrl(scriptUri, templateUrl, mmatch(1),makeAbsolut) & delim, text)
			Next
		Next		

	
		Dim matches4
		KT_preg_match "<param\s+name=(\'|"")?(movie|src|fileName)(\'|"")?([^>]*)>", text, matches
		For each match in matches
			t = match.Value
			t4 = KT_preg_replace ("<param\s+name=(\'|"")?(movie|src|fileName)(\'|"")?([^>]*)>", "$4", t)
			KT_preg_match "\s(value)\s*=\s*((""((\\\""|[^""])+)"")|(\'((\\\\\'|[^\'])+)\')|([^\s]+))?", t4, matches4
			For each match4 in matches4
				tt = match4.Value
				tt4 = KT_preg_replace ("\s(value)\s*=\s*((""((\\\""|[^""])+)"")|(\'((\\\\\'|[^\'])+)\')|([^\s]+))?", "$1^^^$2", tt)
				mmatch = split(tt4, "^^^")
				delim = ""
				If left(mmatch(1),1) = "'" Then
					delim = "'"
					mmatch(1) = mid(mmatch(1), 2, len(mmatch(1))-2)
				ElseIf left(mmatch(1),1) = """" Then
					delim = """"
					mmatch(1) = mid(mmatch(1), 2, len(mmatch(1))-2)
				End If 
				text = KT_preg_replace(mmatch(0) & "\s*=\s*" & KT_preg_quote(delim & mmatch(1) & delim),  mmatch(0) & "=" & delim & KT_Rel2AbsUrl(scriptUri, templateUrl, mmatch(1), makeAbsolut) & delim, text)
			Next
		Next		

		If KT_preg_test("UNI_navigateCancel", text) Then
			KT_preg_match "UNI_navigateCancel\(event, '([\.\/]*includes\/nxt\/back.asp)'\)", text, matches
			For each match in matches
				t = match.Value
				t2 = KT_preg_replace ("UNI_navigateCancel\(event, '([\.\/]*includes\/nxt\/back.asp)'\)", "$1", t)
				text = KT_preg_replace("UNI_navigateCancel\(event, '([\.\/]*includes\/nxt\/back.asp)'\)", "UNI_navigateCancel(event, '" & KT_Rel2AbsUrl(scriptUri, templateUrl, t2, makeAbsolut) & "')", text)
			Next		
		End If

		If KT_preg_test("NEXT_ROOT=", text) Then
			KT_preg_match "NEXT_ROOT=\""([^\""]*)\""", text, matches
			For each match in matches
				t = match.Value
				t2 = KT_preg_replace ("NEXT_ROOT=\""([^\""]*)\""", "$1", t)
				text = KT_preg_replace("NEXT_ROOT=\""([^\""]*)\""", "NEXT_ROOT=""" & KT_Rel2AbsUrl(scriptUri, templateUrl, t2, makeAbsolut) & """", text)
			Next		
		End If

		If KT_preg_test("\$MXW_relPath\s*=\s*", text) Then
			KT_preg_match "\$MXW_relPath\s*=\s*'([^']*)'", text, matches
			For each match in matches
				t = match.Value
				t2 = KT_preg_replace ("\$MXW_relPath\s*=\s*'([^']*)'", "$1", t)
				text = KT_preg_replace("\$MXW_relPath\s*=\s*'([^']*)'", "$MXW_relPath = '" & KT_Rel2AbsUrl(scriptUri, templateUrl, t2, makeAbsolut) & "'", text)
			Next		
		End If

		If KT_preg_test("\""path\"": \"".*includes\/ktm\/\""", text) Then
			KT_preg_match "\""path\"": \""(.*includes\/ktm\/)\""", text, matches
			For each match in matches
				t = match.Value
				t2 = KT_preg_replace ("\""path\"": \""(.*includes\/ktm\/)\""", "$1", t)
				text = KT_preg_replace("\""path\"": \""(.*includes\/ktm\/)\""", """path"": """& KT_Rel2AbsUrl(scriptUri, templateUrl, t2, makeAbsolut) &"""", text)
			Next		
		End If
		
		If InStr(text, "AC_FL_RunContent(") > 0 Then
			KT_preg_match "AC_FL_RunContent\((.*)\);", text, matches
			For each match in matches
				t = match.Value
				t2 = KT_preg_replace("AC_FL_RunContent\((.*)\);", "$1", t)
				
				arr = Split(t2, ",")
				' check for , in the file name
				If (Mid(arr(11), 1, 1) <> "'") Then
					arr(10) = arr(10) & "," & arr(11)
					
					For i = 0 to 10
						auxArr(i) = arr(i)
					Next
					For i = 12 to UBound(arr)
						auxArr(i-1) = arr(i)
					Next
					
					arr = auxArr
				End If
				If UBound(arr) = 17 Then
					If Mid(arr(17), 1, 1) <> "'" Then
						arr(16) = arr(16) & "," & arr(17)
						
						For i = 0 to 16
							auxArr1(i) = arr(i)
						Next
						
						arr = auxArr1
					End If
				End If
				
				arr(10) = Trim(arr(10))
				arr(10) = "'" & KT_Rel2AbsUrl(scriptUri, templateUrl, Mid(arr(10), 2, Len(arr(10)) - 2), makeAbsolut) & "'"
				
				arr(16) = Trim(arr(16))
				arr(16) = "'" & KT_Rel2AbsUrl(scriptUri, templateUrl, Mid(arr(16), 2, Len(arr(16)) - 2), makeAbsolut) & "'"
			
				text = Replace(text, t2, Join(arr, ","))
			Next
		End If
					
		KT_transformsPaths = text
	End Function

	Function KT_Rel2AbsUrl(pageUrl, templateUrl, relUrl, makeAbsolut)
        kt_server = KT_getServerName
        
        If makeAbsolut = false Then
            pageUrl = Replace(pageUrl, kt_server, "")
            templateUrl = Replace(templateUrl, kt_server, "")
            templateUrl = Replace(templateUrl, pageUrl, "")
            kt_server=""
        End If

        If left(relUrl, 1) = "/" Then
			KT_Rel2AbsUrl = kt_server & relUrl
			Exit Function
		End If
		If instr(relUrl, "://") <> 0 Then
			KT_Rel2AbsUrl  = relUrl
			Exit Function
		End If
		
		arrTemplateUrl = split(templateUrl, "/")
		If ubound(arrTemplateUrl) > 0 Then
			Redim Preserve arrTemplateUrl(ubound(arrTemplateUrl)-1)
		Else
			arrTemplateUrl = array()
		End If
		
		Dim ret
		If instr(templateUrl, "://") <> 0 Then
			ret = join(arrTemplateUrl, "/")  & "/" & relUrl
		Else
			arrPageUrl = split (pageUrl, "/")
			If ubound(arrPageUrl) > 0 Then
				Redim Preserve arrPageUrl (ubound(arrPageUrl) - 1)
			Else
				arrPageUrl  = array()
			End If	
			slash = ""
			If ubound(arrTemplateUrl) <> -1 Then
				slash = "/"
			End If
			ret = join (arrPageUrl, "/") & "/"  & join (arrTemplateUrl, "/") & slash & relUrl		
		End If

		KT_Rel2AbsUrl = KT_CanonizeRelPath(ret)
	End Function

	
	Function KT_CanonizeRelPath(relPath__param)
		Dim ret: ret = relPath__param

        Do while True
            ' remove the ./
            found = false
            
            pos = instr(ret, "./")
            If pos <> 0 Then
                left_part = left(ret, pos-1)
                right_part = mid(ret, pos + 2)
                If left_part = "" Or right(left_part, 1) = "/" Then
                    ret = left_part & right_part
                    found = true
                End If	
            End If
            
            If Not found Then
                pos = instr(ret, "/..")
                If pos <> 0 Then
                    left_part = left(ret, pos - 1)
    
                    If left_part = ".." Then
                        Exit Do
                    End If
                    
                    right_part = mid(ret, pos + 3)
                    
                    last_delim = instrrev(left_part, "/")
                    If last_delim <> 0 Then
                        left_part = left (left_part, last_delim - 1)
                    Else
                        left_part = ""
                        If left(right_part, 1) = "/" Then
                            right_part = right (right_part, len(right_part)-1)
                        End If	
                    End If
                    ret = left_part & right_part
                    
                    found = true				
                End If	
            End If
            
			If Not found Then
				Exit Do	
			End If	
		Loop
		
		KT_CanonizeRelPath = ret
	End Function



	Function KT_makeIncludedURL (url) 
		dim ret
		ret = url
		If KT_REL_PATH <> "" Then
			If Not KT_preg_test("^/", ret) And Not KT_preg_tesT("^[a-z]+://", ret) Then
				ret = KT_REL_PATH & ret
			End If
		End If
		KT_makeIncludedURL = ret
	End Function

	Function KT_makeIncludedPath (path) 
		Dim ret
		ret = path
		If Not KT_preg_test ("^[A-Za-z]:(\\[^\\:\*\?\""<>|]*)*$", path) And Not KT_preg_test("^/", path) Then
			If KT_REL_PATH <> "" Then
				ret = KT_REL_PATH & ret
			End If			
		End If
		KT_makeIncludedPath = ret
	End Function

	Public Function KT_pathinfo(path)
		Set pathinfo = Server.CreateObject("Scripting.Dictionary")
		escaped_path = replace(path, "/", "\")
		pos = instrrev(escaped_path, "\")
		If pos <> 0 Then
			dirname = left (path, pos)
			fullfilename = mid (path, pos + 1) 
		Else
			fullfilename = path	
		End If

		filename = fullfilename
		extension = ""
		dot = ""
		pos = instrrev(fullfilename, ".")
		If pos  <> 0 Then
			dot = "."
			extension = mid(fullfilename, pos+1)
			filename = mid(fullfilename, 1, pos-1)
		End If
		pathinfo("dot") = dot
		pathinfo("filename") = filename
		pathinfo("dirname") = dirname
		pathinfo("extension") = extension
		pathinfo("basename") = filename & dot & extension
		Set KT_pathinfo = pathinfo
	End Function

	
	Public Function KT_dirname(path)
		Set path_info = KT_pathinfo(path)
		KT_dirname = path_info("dirname")
	End Function


	Sub KT_setDbType(ByRef connection)
		dbType = ""
		If isobject(connection) Then
			conn_str = connection.ConnectionString
			If instr(1, conn_str, "Microsoft.Jet.OLEDB", 1) <> 0 Then
				dbType = "access_oledb"
			End If
			If instr(1, conn_str, "mysql", 1) <> 0 Then	
				dbType = "mysql"
			End If
			If instr(1, conn_str, "option=",1) <> 0 Then 
				' in order to cover OLEDB PRIVIDER FOR ODBC Drivers
				dbType = "mysql"
			End If
		End If
		ExecuteGlobal "Dim KT_DatabaseType:  KT_DatabaseType = """ & dbType & """"
	End Sub


	'
	' Escapes a value against a specific type to be used in the transaction SQL
	' Ex: value=ab'b and type=STRING, result=ab\'b (escapes slashes)
	' @param object unknown $colValue The value to prepare
	' @param string $colType The type (STRING_TYPE, NUMERIC_TYPE, etc)
	' @return object unknown The escaped value
	' @access public
	'
	Function KT_escapeForSql(colValue__param, colType__param)
			Dim colValue, colType
			colValue = colValue__param
			colType = colType__param
			' initialisation of all properties goes here
			Set type2empty  = Server.CreateObject("Scripting.Dictionary")
			type2empty("STRING_TYPE")		=	 "null"
			type2empty("NUMERIC_TYPE") 		=	 "null"
			type2empty("DOUBLE_TYPE") 		=	 "null"
			type2empty("DATE_TYPE")			=	 "null"
			type2empty("DATE_ACCESS_TYPE")	=	 "null"
			type2empty("FILE_TYPE")			=	 "null"
			type2empty("CHECKBOX_YN_TYPE")	=	 "'N'"
			type2empty("CHECKBOX_1_0_TYPE")	=	 "0"
			type2empty("CHECKBOX_-1_0_TYPE")=	 "0"
			type2empty("CHECKBOX_TF_TYPE")	=	 "'f'"
			' correspondence between a datatype and the quote
			set type2quote = Server.CreateObject("Scripting.Dictionary")
			type2quote("STRING_TYPE")		=	 "'"
			type2quote("NUMERIC_TYPE") 		=	 ""
			type2quote("DOUBLE_TYPE") 		=	 ""		
			type2quote("DATE_TYPE")			=	 "'"
			type2quote("DATE_ACCESS_TYPE")	=	 "#"
			type2quote("FILE_TYPE")			=	 "'"
			type2quote("CHECKBOX_YN_TYPE")	=	 "'"
			type2quote("CHECKBOX_1_0_TYPE")	=	 ""
			type2quote("CHECKBOX_-1_0_TYPE")=	 ""
			type2quote("CHECKBOX_TF_TYPE")	=	 "'"
			
			If isEmpty(colValue) Or isNull(colValue) Or colValue = "" Then
				tmValue = type2empty(colType)	
			Else
				On Error Resume Next
				If colType = "NUMERIC_TYPE" Then
					colValue = Clng(colValue)
				ElseIf colType = "DOUBLE_TYPE" Then
					colValue = CDbl(replace(colValue, ",", ".")) ' colValue = replace(CDbl(colValue), ",", ".")
				End If
				If Err.Number <> 0 Then
					colValue = 0  ' init to a default value
				End If
				On Error GoTo 0
				
				colValue = replace (colValue, "'","''")
				If KT_DatabaseType = "mysql" Then
					colValue = replace (colValue, "\","\\")
				End If
				quote = type2quote(colType)
				tmValue = quote & colValue & quote
			End If
			KT_escapeForSql = tmValue
	End Function


	Function KT_escapeForFakeRs(colValue, colType)
			' initialisation of all properties goes here
			Set type2empty  = Server.CreateObject("Scripting.Dictionary")
			type2empty("STRING_TYPE")		=	 ""
			type2empty("NUMERIC_TYPE") 		=	 ""
			type2empty("DOUBLE_TYPE") 		=	 ""
			type2empty("DATE_TYPE")			=	 ""
			type2empty("DATE_ACCESS_TYPE")	=	 ""
			type2empty("FILE_TYPE")			=	 ""
			type2empty("CHECKBOX_YN_TYPE")	=	 "N"
			type2empty("CHECKBOX_1_0_TYPE")	=	 "0"
			type2empty("CHECKBOX_-1_0_TYPE")=	 "0"
			type2empty("CHECKBOX_TF_TYPE")	=	 "f"
			
	
			If isEmpty(colValue) Or isNull(colValue) Or colValue = "" Then
				tmValue = type2empty(colType)	
			Else
				On Error Resume Next
				If colType = "NUMERIC_TYPE" Then
					colValue = Clng(colValue)
				ElseIf colType = "DOUBLE_TYPE" Then
					colValue = CDbl(replace(colValue, ",", "."))
				End If
				If Err.Number <> 0 Then
					colValue = 0  ' init to a default value
				End If
				On Error GoTo 0
				
				tmValue = colValue
			End If
			KT_escapeForFakeRs = tmValue
	End Function

	
	'
	' Escapes a value of a field name to be used in the transaction SQL
	'  Ex: First Name gets translated into `First Name`
	'  @param string $colName The DataBase field name
	'  @return string The escaped field name
	'  @access public
	'
	Function KT_escapeFieldName(colName)
		KT_escapeFieldName = colName
		Exit Function
		
		mysql_start_quote = "`"
		mysql_end_quote = "`"

		access_oledb_start_quote = "["
		access_oledb_end_quote = "]"
		
		others_start_quote = """"
		others_end_quote = """"
		
		
		If KT_DatabaseType = "mysql" Then
			start_quote = mysql_start_quote
			end_quote = mysql_end_quote
		ElseIf KT_DatabaseType = "access_oledb" Then
			start_quote = access_oledb_start_quote
			end_quote = access_oledb_end_quote
		Else
			start_quote = others_start_quote
			end_quote = others_end_quote
		End If


		If Instr(colName, ".") <> 0 Then
			KT_escapeFieldName = colName
			Exit Function
		End If
		
		If KT_preg_test("^`(.+)`$", colName) Then
			colName = KT_preg_replace("^`(.+)`$", "$1", colName)
		End If
		
		KT_escapeFieldName = start_quote & colName & end_quote
	End Function
	
	' computes the number of rows in the recordset, and then resets recordset
	Function KT_getNumberOfRows(ByRef rec)
		Dim rec_total
		
		' set the record count
		rec_total = rec.RecordCount
		If (rec_total = -1) Then
		  ' count the total records by iterating through the recordset
		  rec_total=0
		  While (Not rec.EOF)
			rec_total = rec_total + 1
			rec.MoveNext
		  Wend
		
		  ' reset the cursor to the beginning
		  If (rec.CursorType > 0) Then
			rec.MoveFirst
		  Else
			rec.Requery
		  End If
		End If
		
		KT_getNumberOfRows = rec_total
	End Function


	Function KT_min(number1, number2)
		If number1 < number2 Then
			KT_min = number1
		Else
			KT_min = number2		
		End If
	End Function
	
	Function KT_max(number1, number2)
		If number1 > number2 Then
			KT_max = number1
		Else
			KT_max = number2		
		End If
	End Function
	
	Function KT_SessionKtBack(newUrl__param)
		newUrl = newUrl__param
		If Not KT_isSet(Session("KT_backArr")) Then
			Session("KT_backArr") = Array(newUrl)
		ElseIf	Ubound(Session("KT_backArr")) <> -1 Then
		 	' pop it up
			KT_backArr	= Session("KT_backArr")
			oldUrl = KT_backArr(Ubound(KT_backArr))
			If Ubound(KT_backArr) = 0 Then
				KT_backArr = Array()
			Else
				Redim Preserve KT_backArr(Ubound(KT_backArr)-1)
			End If
			saveOld = oldUrl
			saveNew = newUrl
			
			toldUrl = KT_addReplaceParam(oldUrl,  "#pageNum_.*", "")
			toldUrl = KT_addReplaceParam(toldUrl, "#sorter_.*", "")
			toldUrl = KT_addReplaceParam(toldUrl, "#show_all_.*", "")
			toldUrl = KT_addReplaceParam(toldUrl, "#show_filter_.*", "")
			
			tnewUrl = KT_addReplaceParam(newUrl,  "#pageNum_.*", "")
			tnewUrl = KT_addReplaceParam(tnewUrl, "#sorter_.*", "")
			tnewUrl = KT_addReplaceParam(tnewUrl, "#show_all_.*", "")
			tnewUrl = KT_addReplaceParam(tnewUrl, "#show_filter_.*", "")

			If tnewUrl <> toldUrl Then
				KT_backArr = KT_array_push(KT_backArr, saveOld)
			End If
			KT_backArr = KT_array_push(KT_backArr, saveNew)		
			Session("KT_backArr") = KT_backArr			
		Else
			KT_backArr	= Session("KT_backArr")
			KT_backArr	= KT_array_push(KT_backArr, newUrl)
			Session("KT_backArr") = KT_backArr
		End If
	End Function

	Function KT_GetPooledConnection(ConnectionString__param)
		pooledConnectionString = replace(ConnectionString__param  & "", """", """""")
		strToExecute = "" & _
		"If isEmpty(KT_ConnectionsPool) Then" & vbNewLine & _
		"	Set KT_ConnectionsPool = Server.CreateObject(""Scripting.Dictionary"")" & vbNewLine & _
		"End If" & vbNewLine & _
		"If KT_ConnectionsPool.Exists(""" & pooledConnectionString & """) Then" & vbNewLine & _
		"	Set KT_GlobalConnection = KT_ConnectionsPool(""" & pooledConnectionString & """)"  & vbNewLine & _
		"Else"  & vbNewLine & _
		"	Set KT_GlobalConnection = Server.CreateObject(""ADODB.Connection"")" & vbNewLine & _
		"	KT_GlobalConnection.ConnectionString = """ & pooledConnectionString & """" & vbNewLine & _
		"	KT_GlobalConnection.Open"  & vbNewLine & _
		"	Set KT_ConnectionsPool(""" & pooledConnectionString & """) = KT_GlobalConnection"  & vbNewLine & _	
		"End If"
		
		ExecuteGlobal strToExecute
		Set KT_GetPooledConnection  = KT_GlobalConnection
	End Function
	
	Function KT_Server_MapPath(path)
		If KT_preg_test ("^[A-Za-z]:(\\[^\\:\*\?\""<>|]*)*$", path) Or KT_preg_test ("^\\\\([A-Za-z_0-9]+|(\d{1,3}\.){3}\d{1,3})(\\[^\\:\*\?\""<>|]*)*$", path) Then
			KT_Server_MapPath = path
		Else
			If left(path, 1) = "/" Or left(path, 1) = "\" Then
				KT_Server_MapPath = Server.MapPath(path)
			Else
				KT_Server_MapPath = KT_normalizeAbsoluteFilePath(Server.MapPath(".") & "\" & path)
			End If
		End If
	End Function
	
'
' Function KT_DynamicData
' @param string $expression The expression to be evaluated
' @param object or null $tNG The tNG context in which the expression is evaluated
' @param string $escapeMethod The string escape method for the evaluated values (rawurlencode and SQL)
' @param booolean $useSavedData Weather to use the current tNG data or the saved values
' @param array $extraParams Extra expression parameters passed when for evaluation (of form $key => $value; any encounter of key will be replaced with its value)
'
Function KT_DynamicData(expression, ByRef tNG, escapeMethod, useSavedData, extraParams, errorIfNotFound)
	KT_DynamicData = expression
	If isEmpty(expression) Or isnull(expression) Or expression = "" Then
		Exit Function
	End If

	If isNull(escapeMethod) Then
		escapeMethod = ""
	End If
	If isNull(useSavedData) Then
		useSavedData = false
	End If
	If Not KT_isSet(extraParams) Then
		Set extraParams = Server.CreateObject("Scripting.Dictionary")
	End If
	If isNull(errorIfNotFound) Then
		errorIfNotFound = true
	End If
	
	PB = "{"
	PE = "}"
		
	Set placeholdersArr = KT_getReplacementsFromMessage(expression)
	dynamicDataFunctions = Array("NOW()", "now()", "NOW", "now", _
							"NOW_DT()", "now_dt()", "NOW_DT", "now_dt", _
							"NOW_T()", "now_t()", "NOW_T", "now_t", _
							"KT_REFERRER", "kt_referrer", "KT_CSV_LINE", "KT_XML_LINE")
	Set replacementsArr = Server.CreateObject("Scripting.Dictionary")
	Select Case escapeMethod
		Case "rawurlencode"
		
		Case "expression"
		
		Case "SQL"
			If not KT_isSet(tNG) Then
				escapeMethod = false
			End If
		Case Else
			escapeMethod = false	
	End Select

	For each idx in placeholdersArr
		placeholder = placeholdersArr(idx)
		If KT_in_array(placeholder, KT_array_keys(extraParams), true) Then
			'  extra params have priority 1
			placeholderType = "tNG_DDextra"
			placeholderName = placeholder
		Else
			' functions have priority 2
			If KT_in_array(placeholder, dynamicDataFunctions, true) Then
				placeholderType = "tNG_DDfunction"
				placeholderName = placeholder
			Else
				ptpos = instr(placeholder, ".")
				If ptpos = 0 Then
					' tNG field
					If KT_isSet(tNG) Then
						' attached to a tng, replace field with value
						placeholderType = "tNG_tNGfield"
					 	placeholderName = placeholder
					Else
						' no tng, leave as is
						placeholderType = "tNG_tNGfieldLater"
					 	placeholderName = placeholder
					End If
				 Else	
					placeholderType = left(placeholder, ptpos-1)
					placeholderName = mid(placeholder, ptpos+1)
				 End If
			End If
		End If

		placeholder = PB & placeholder & PE
		Select case LCase(placeholderType)
			case "tng_ddfunction"
				replacementsArr(placeholder) = KT_GetDynamicDataFunctionValue(placeholderName)
			case "tng_ddextra"
				replacementsArr(placeholder) = extraParams(placeholderName)
			case "tng_tngfield"
				If useSavedData Then
					placeholderValue = tNG.getSavedValue(placeholderName)
				Else
					Set cols = tNG.columns
					If cols.Exists(placeholderName) Or tNG.getPrimaryKey() = placeholderName Then
						placeholderValue = tNG.getColumnValue(placeholderName)
						placeholderType = tNG.getColumnType(placeholderName)
					Else
						If errorIfNotFound Then
							Response.write 	"KT_DynamicData:<br />Column " & placeholderName & " is not part of the current transaction."
							Response.End()
						Else
							placeholderValue = placeholder
						End If
					End If
					If escapeMethod = "SQL" Then
						placeholderValue = KT_escapeForSql(placeholderValue, placeholderType)
					End If
				End If
				replacementsArr(placeholder) = placeholderValue
			
			case "tng_tngfieldlater"

			case "get"
				myPlaceholderName = placeholderName
				If KT_isSet(tNG) Then
					If KT_isSet(tNG.multipleIdx) Then
						myPlaceholderName = myPlaceholderName & "_" & tNG.multipleIdx
					End If
				End If
				replacementsArr(placeholder) = KT_getRealValue("GET", myPlaceholderName)
				If Not KT_isSet(replacementsArr(placeholder)) Then
					replacementsArr(placeholder) = KT_getRealValue("GET", placeholderName)
				End If
				
			case "post"
				myPlaceholderName = placeholderName
				If KT_isSet(tNG) Then
					If KT_isSet(tNG.multipleIdx) Then
						myPlaceholderName = myPlaceholderName & "_" & tNG.multipleIdx
					End If
				End If
				replacementsArr(placeholder) = KT_getRealValue("POST", myPlaceholderName)
				If Not KT_isSet(replacementsArr(placeholder)) Then
					replacementsArr(placeholder) = KT_getRealValue("POST", placeholderName)					
				End If

			case "request"
				myPlaceholderName = placeholderName
				If KT_isSet(tNG) Then
					If KT_isSet(tNG.multipleIdx) Then
						myPlaceholderName = myPlaceholderName & "_" & tNG.multipleIdx
					End If
				End If
				replacementsArr(placeholder) = KT_getRealValue("REQUEST", myPlaceholderName)
				If Not KT_isSet(replacementsArr(placeholder)) Then
					replacementsArr(placeholder) = KT_getRealValue("REQUEST", placeholderName)
				End If
				
			case "server"
				replacementsArr(placeholder) = KT_getRealValue("SERVER", placeholderName)
				
			case "cookie"
				replacementsArr(placeholder) = KT_getRealValue("COOKIE", placeholderName)
				
			case "session"
				replacementsArr(placeholder) = KT_getRealValue("SESSION", placeholderName)

			case "application"
				replacementsArr(placeholder) = KT_getRealValue("APPLICATION", placeholderName)
				
			case "globals"
				replacementsArr(placeholder) = KT_getRealValue("GLOBALS", placeholderName)

			case "csv"
				replacementsArr(placeholder) = KT_getRealValue("CSV", placeholderName)
				
			case "xml"
				replacementsArr(placeholder) = KT_getRealValue("XML", placeholderName)
				
			case else
				' recordset
				On Error Resume Next
				str = "Set KT_rsDD = " & placeholderType & vbNewLine & _
					"KT_valDD = KT_rsDD.Fields.Item(""" & placeholderName & """).Value"
				ExecuteGlobal str
				If Err.Number <> 0 Then
					placeholderValue = placeholder
				Else
					placeholderValue = KT_valDD
				End If
				On Error GoTo 0
				replacementsArr(placeholder) = placeholderValue
				
		End Select
	Next

	If escapeMethod = "rawurlencode" Then
		If Not replacementsArr.Exists("{kt_login_redirect}")  And Not replacementsArr.Exists("{kt_referrer}") And Not replacementsArr.Exists("{KT_REFERRER}") Then
			For each k in replacementsArr
				replacementsArr(k) = Server.URLEncode(replacementsArr(k) & "")
			Next	
		End If
	ElseIf escapeMethod = "expression" Then
		For each k in replacementsArr
			replacementsArr(k) = KT_escapeExpression(replacementsArr(k))
		Next	
	End If	
	' build the expression
	newexpression = expression
	For each k in replacementsArr
		newexpression = replace(newexpression, k, replacementsArr(k) & "")
	Next
	KT_DynamicData = newexpression
End Function	
	
	
Function KT_getReplacementsFromMessage(str) 
	Dim ret: Set ret = Server.CreateObject("Scripting.Dictionary")
	Dim matches
	Dim indexmatches: indexmatches = 0
	KT_preg_match "\{([\w\d\.\s\(\)]+)\}", str, matches
	For each match in matches
		ret(indexmatches) = replace(replace(match.Value, "}", ""), "{", "")
		indexmatches = indexmatches + 1
	Next
	Set KT_getReplacementsFromMessage = ret
End Function

Function KT_GetDynamicDataFunctionValue( placeholder )
	KT_GetDynamicDataFunctionValue = placeholder

	KT_getInternalTimeFormat()
	' getting date	
	Dim yyyy: yyyy = CInt(DatePart("yyyy", Date()))
	Dim m: m = CInt(DatePart("m",  Date()))
	If m < 10 Then
		m = "0" & m
	End If
	Dim d: d = CInt(DatePart("d",  Date()))
	If d < 10 Then
		d = "0" & d
	End If
	current_date = yyyy & "-" & m & "-" & d
	
	
	' getting time
	Dim h: h = Cint(Hour(Now))
	If h < 10 Then
		h = "0" & h
	End If
	Dim min: min = Cint(Minute(Now))
	If min < 10 Then
		min = "0" & min
	End If
	Dim sec: sec = Cint(Second(Now))
	If sec < 10 Then
		sec = "0" & sec
	End If	
	current_time = h & ":" & min & ":" & sec
	
	If KT_in_array(placeholder, Array("NOW()", "now()", "NOW", "now"), true) Then
		date_now = KT_convertDate(current_date, "yyyy-mm-dd", KT_screen_date_format)	
		KT_GetDynamicDataFunctionValue = date_now
	ElseIf KT_in_array(placeholder, Array("NOW_DT()", "now_dt()", "NOW_DT", "now_dt"), true) Then
		date_dt_now = KT_convertDate(current_date & " " & current_time, "yyyy-mm-dd HH:ii:ss", KT_screen_date_format & " " & KT_screen_time_format_internal)	
		KT_GetDynamicDataFunctionValue = date_dt_now
	ElseIf KT_in_array(placeholder, Array("NOW_T()", "now_t()", "NOW_T", "now_t"), true) Then
		date_t_now = KT_convertDate(current_time, "HH:ii:ss", KT_screen_time_format_internal)
		KT_GetDynamicDataFunctionValue = date_t_now
	ElseIf KT_in_array(placeholder, Array("KT_REFERRER", "kt_referrer"), true) Then
		KT_referrer = Request.ServerVariables("HTTP_REFERER")
		KT_GetDynamicDataFunctionValue = KT_referrer
	ElseIf placeholder = "KT_CSV_LINE" Then
		KT_GetDynamicDataFunctionValue = KT_CSV_LINE
	ElseIf placeholder = "KT_XML_LINE" Then
		KT_GetDynamicDataFunctionValue = KT_XML_LINE
	End If	
End Function	

'
' Compiles a method and a reference into a value
' Ex: method=GET and reference=test, return value=$_GET['test']
' @param string $method The method (GET, POST, etc)
' @param string $reference The reference (variable name)
' @return object unknown The compiled value
'         the return has stripped slashes if necessary
'

Public Function KT_getRealValue (method, reference) 	
	ret	 = null
	Select case method
		case "GET"
			If KT_isSet( Request.QueryString(reference)) Then	
				ret = Request.QueryString(reference)
			End If
			
		case "POST"
			If KT_isSet( Request.Form(reference)) Then	
				ret = Request.Form(reference)
			End If

		case "REQUEST"
			If KT_isSet( Request(reference)) Then	
				ret = Request(reference)
			End If
		
		case "COOKIE"
			If KT_isSet( Request.Cookies(reference)) Then	
				ret = Request.Cookies(reference)
			End If

		case "SESSION"
			If KT_isSet(Session(reference)) Then	
				ret = Session(reference)
			End If
			
		case "APPLICATION"
			If KT_isSet(Application(reference)) Then	
				ret = Application(reference)
			End If
						
		case "GLOBALS"
			ExecuteGlobal "MYGLOBAL = " & reference
			ret = MYGLOBAL
			
		case "SERVER":
			If KT_isSet(Request.ServerVariables(reference)) Then
				ret = Request.ServerVariables(reference)
			End If	

		case "FILES"
			ret = Request.Form(reference)


		case "VALUE"
			ret = reference

		case "CURRVAL"
			ret = null
			
		case "CSV"
			On Error Resume Next
			If KT_CSV.Exists(reference) Then
				ret = KT_CSV(reference)
			End If	
			On Error GoTo 0
			
		case "XML"
			On Error Resume Next
			If KT_XML.Exists(reference) Then
				ret = KT_XML(reference)
			End If	
			On Error GoTo 0
			
		case else
			Response.write "KT_getRealValue:<br />Unknown method: " & method & "."
			Response.End()

	End Select
	KT_getRealValue = ret
End	Function
	

Function KT_normalizeAbsolutePath (path)
    KT_createGlobalFSO
	KT_normalizeAbsolutePath = KT_FSO.GetAbsolutePathName(path & "\") & "\"
End Function

	
Function KT_normalizeAbsoluteFilePath (path)
	KT_createGlobalFSO
	KT_normalizeAbsoluteFilePath = KT_FSO.GetAbsolutePathName(path)
End Function

Function KT_addSuffixToPath(path, suffix)
	Dim ret
	ret = path
	If ret <> "" Then
		If right(ret, 1) <> "\" And right(ret, 1) <> "/" Then
			ret  = ret & suffix
		End If
	End If
	KT_addSuffixToPath = ret	
End Function

' Leave just ASCII chars (alpha-numeric) and _ . , ; () [] 
' @param string string to be parsed;
' @param string if equal folder let /;
' @return string;
' @access public
Function KT_replaceSpecialChars(text, what)
	If what = "folder" Then
		KT_replaceSpecialChars = Trim(KT_preg_replace("[^\/0-9a-z\.,;\-_\(\)\[\]\s]", "_", text))
	Else
		KT_replaceSpecialChars = Trim(KT_preg_replace("[^0-9a-z\.,;\-_\(\)\[\]\s]", "_", text))
	End If
End Function

Function KT_functionExists(ByVal functionName)
  Dim obj
  On Error Resume Next
    Set obj = GetRef(functionName)
    If Err.Number <> 0 Then
      KT_functionExists = False
    Else
      KT_functionExists = True
    End If
  On Error GoTo 0
End Function

Function KT_cleanContent (value_param, allowedTags_param, deniedTags_param)
	Dim arrAllowed: arrAllowed = Split(Trim(allowedTags_param),",",-1,1)
	Dim arrDenied: arrDenied = Split(Trim(deniedTags_param),",",-1,1)
	If Trim(deniedTags_param) <> "" AND Ubound(arrDenied) >= 0 Then
		KT_preg_match "<(/?)(\w+)[^>]*>", value_param, matches
			For each match in matches
				t = match.Value
				tag_name = KT_preg_replace ("<(/?)(\w+)[^>]*>", "$2", t)
				tag_name = lcase(tag_name)

				If KT_in_array(tag_name, arrDenied, false) Then
					value_param = replace(value_param, t, "")
				End If		
			Next
	Else
		If Trim(allowedTags_param) <> "" AND Ubound(arrAllowed) >= 0  Then
			KT_preg_match "<(/?)(\w+)[^>]*>", value_param, matches
			For each match in matches
				t = match.Value
				tag_name = KT_preg_replace ("<(/?)(\w+)[^>]*>", "$2", t)
				tag_name = lcase(tag_name)
				If Not KT_in_array(tag_name, arrAllowed, false) Then
					value_param = replace(value_param, t, "")
				End If
			Next
		End If
	End If
	
	
	
	' remove script tag
	If KT_in_array("script",arrDenied,false) OR Ubound(arrAllowed) > 0 AND NOT KT_in_array("script",arrAllowed,false) Then

		'remove on... attributes 
		KT_preg_match "<[^/]([^>]+)>",value_param,matches
		For Each match in matches
			t = match.Value
			KT_preg_match "(on[a-z]+\s*=\s*""[\w\W]*?"")|(on[a-z]+\s*=\s*'[\w\W]*?')|(on[a-z]+\s*=\s*[^\s]*)", t,mmatches
			For each mmatch in mmatches
				t2 = mmatch.Value
				value_param = replace(value_param,t2,"")
			Next
		Next

	' remove javascript: in tags
		KT_preg_match "<[^/]([^>]+)>",value_param,matches
		For Each match in matches
			t = match.Value
			KT_preg_match "(\W\w+\s*=\s*\""([^\""]*j\s*a\s*v\s*a\s*s\s*c\s*r\s*i\s*p\s*t\s*:[^\""]*)\""+)|(\W\w+\s*=\s*'([^']*j\s*a\s*v\s*a\s*s\s*c\s*r\s*i\s*p\s*t\s*:[^']*)'+)|(\W\w+\s*=\s*j\s*a\s*v\s*a\s*s\s*c\s*r\s*i\s*p\s*t\s*[^\s>]*)", t,mmatches
			For each mmatch in mmatches
				t2 = mmatch.Value
				value_param = replace(value_param,t2,"")
			Next
		Next
	

'remove :expression from style
		KT_preg_match "<\s*style[\w\W]*>([\s\S]*)<\/style>",value_param,matches
		For Each match in matches
			t = match.Value
			KT_preg_match ":\s*(expression)\s*\(", t,mmatches
			For each mmatch in mmatches
				t2 = mmatch.Value
				value_param = replace(value_param,t2,":")
			Next
		Next
	
	'remove :expression from style
		KT_preg_match "<(.*)style\s*=[^>]+>",value_param,matches
		For Each match in matches
			t = match.Value
			KT_preg_match ":\s*(expression)\s*\(", t,mmatches
			For each mmatch in mmatches
				t2 = mmatch.Value
				value_param = replace(value_param,t2,":")
			Next
		Next
	End If
	KT_cleanContent = value_param
End Function
%>
