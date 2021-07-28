<%
'**********************************
' http://www.DwZone-it.com
' Random Recordset
' Copyright (c) DwZone.it 2000-2005
'**********************************

class dwzRecPaging
	
	Private dwzRec
	Private Recordset
	private PagingType
	Private RecPaging
	Private Pages
	Private sNext
	Private sPrevious
	Private sFirst
	Private sLast
	Private StyleLink
	Private CurStyleLink
	Private InactiveStyleLink
	Private Separator
	Private ActiveLinkMask
	Private CurrentPage
	Private MaxRecord
	Private TotalPage
	Private EmptyElements
	Private DisplayNumericEntry
	Private RecField
	Private CurrentLetter
	Private Letters()
	Private Fields()
	Private CurrentField
	Private nFields
	
	Public sub setRecField(param)
		RecField = param
	end sub
	
	Public sub setDisplayNumericEntry(param)
		DisplayNumericEntry = param
	end sub
	
	Public sub setEmptyElements(param)
		EmptyElements = param
	end sub
	
	Public sub setLinkMask(param)
		ActiveLinkMask = param
	end sub
	
	Public Sub setSeparator(param)
		Separator = replace(param, " ", "&nbsp;")
	End Sub	
	
	Public sub setRecordset (ByRef rs)
		Set Recordset = rs
	End Sub
	
	Public Sub setTypeNumeric()
		PagingType = "N"
	End Sub	
	
	Public sub setTypeFieldValue()
		PagingType = "S"
	End Sub	
	
	Public Sub setTypeLetter()
		PagingType = "L"
	End Sub	
	
	Public Sub setRecPaging(param)
		RecPaging = param
	End Sub	
	
	Public Sub setPages(param)
		Pages = param
	End Sub	
		
	Public Sub setStyle(param1, param2, param3)
		if param1 <> "None" then
			StyleLink = param1
		end if
		if param2 <> "None" then
			CurStyleLink = param2
		end if
		if param3 <> "None" then
			InactiveStyleLink = param3
		end if
	End Sub	
		
	Public Sub setText(param1, param2, param3, param4)
		sNext = param1
		sPrevious = param2
		sFirst = param3
		sLast = param4
	End Sub	
		
		
	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
		set dwzRec = nothing
	End Sub
		
	Public Sub Init()
		if request("dwzPage") <> "" then
			CurrentPage = clng(request("dwzPage"))
		else
			CurrentPage = 1
		end if
		if request("dwzLetter") <> "" then
			CurrentLetter = request("dwzLetter")
		else
			CurrentLetter = "A"
		end if
		
		if request("dwzField") <> "" then
			CurrentField = clng(request("dwzField"))
		else
			CurrentField = 0
		end if
		nFields = -1
		Redim Letters(27)
		for J=0 to 27
			Letters(J) = 0
		next
	End Sub	
	
	private function GetNumericPaging()
		retStr = ""
		startLink = GetQueryStringWithPage()
		
		startLink = replace(startLink, "&", "&amp;")
		
		LinkPrevious = ""
		LinkNext = ""
		LinkFirst = ""
		LinkLast = ""
		
		if trim(sFirst) <> "" then
			if CurrentPage <> 1 then
				LinkFirst = "<a href='" & Replace(Replace(startLink, """", "%22"), ",", "") & "dwzPage=1" & "' class='ContentLinks' >" & sFirst & "</a>"
			else
				LinkFirst = "<span class='PagingNonLinks'>" & sFirst & "</span>"
			end if
			if retStr <> "" then
				retStr = retStr & Separator
			end if
			retStr = retStr & LinkFirst
		end if
		
		if trim(sPrevious) <> "" then
			if CurrentPage > 1 then
				LinkPrevious = "<a href='" & Replace(Replace(startLink, """", "%22"), ",", "") & "dwzPage=" & cstr(CurrentPage-1) & "' class='ContentLinks' >" & Replace(Replace(sPrevious, """", "%22"), ",", "") & "</a>"
			else
				LinkPrevious = "<span class='PagingNonLinks'>" & sPrevious & "</span>"
			end if
			if retStr <> "" then
				retStr = retStr & Separator
			end if
			retStr = retStr & LinkPrevious
		end if
		
		if Pages = -1 then
			inizio = 1
			fine = TotalPage
		else
			inizio = CurrentPage - int(Pages / 2)
			if inizio < 1 then
				inizio = 1
			end if
			fine = inizio + pages
			if fine > TotalPage then
				fine = TotalPage
			end if
			if fine - inizio >= pages then
				fine = fine - 1
			elseif fine - inizio < pages then
				inizio = fine - (pages - 1)
				if inizio < 1 then
					inizio = 1
				end if
			end if
		end if	
		
		for J=inizio to fine			
			if J = CurrentPage then
				Link = "<span class='" & CurStyleLink & "'>" & replace(ActiveLinkMask, "{1}", cstr(J), 1, -1, vbtextcompare) & "</span>"
			else
				Link = "<a href='" & Replace(Replace(startLink, """", "%22"), ",", "") & "dwzPage=" & cstr(J) & "' class='ContentLinks' >" & cstr(J) & "</a>"
			end if	
			if retStr <> "" then
				retStr = retStr & Separator
			end if
			retStr = retStr & Link
		next
		
		if trim(sNext) <> "" then
			if CurrentPage < TotalPage then
				LinkNext = "<a href='" & Replace(Replace(startLink, """", "%22"), ",", "") & "dwzPage=" & cstr(CurrentPage + 1) & "' class='ContentLinks' >" & Replace(Replace(sNext, """", "%22"), ",", "") & "</a>"
			else
				LinkNext = "<span class='PagingNonLinks'>" & Replace(sNext, """", "%22") & "</span>"
			end if
			if retStr <> "" then
				retStr = retStr & Separator
			end if
			retStr = retStr & LinkNext
		end if
		
		if trim(sLast) <> "" then
			if CurrentPage <> TotalPage then
				LinkLast = "<a href='" & Replace(Replace(startLink, """", "%22"), ",", "") & "dwzPage=" & cstr(TotalPage) & "' class='ContentLinks' >" & Replace(Replace(sLast, """", "%22"), ",", "") & ": <b>" & cstr(TotalPage) & "</b></a>"
			else
				LinkLast = "<span class='PagingNonLinks'>" & Replace(Replace(sLast, """", "%22"), ",", "") & "</span>"
			end if
			if retStr <> "" then
				retStr = retStr & Separator
			end if
			retStr = retStr & LinkLast
		end if
		GetNumericPaging = retStr
	end function
	


	Public Sub GetPaging()
		if PagingType = "L" then
			response.write GetLetterPaging()
		elseif PagingType = "S" then
			response.write GetStringPaging()
		else
			response.write GetNumericPaging()
		end if
	end sub
			
	private function GetQueryStringWithPage()
		retStr = request.ServerVariables("PATH_INFO")
		qString = ""
		for each key in request.QueryString
			if lcase(key) <> "dwzpage" and lcase(key) <> "dwzletter" and lcase(key) <> "dwzfield" then
				if qString <> "" then
					qString = qString & "&"
				end if
				qString = qString & key & "=" & request.QueryString(key)
			end if
		next
		retStr = retStr & "?"
		if qString <> "" then
			retStr = retStr & qString & "&"		
		end if
		GetQueryStringWithPage = retStr
	end function
	
	Public sub Execute()
		If Not isObject(recordset) Then
			Response.write "<strong>DwZone - Recordset Paging Error.</strong><br/>The recordset is not valid"
			Exit sub
		End If
		
		if recordset.eof then
			set dwzRec = recordset
			exit sub
		end if
		
		Const adEmpty = 0
		Const adTinyInt = 16
		Const adSmallInt = 2
		Const adInteger = 3
		Const adBigInt = 20
		Const adUnsignedTinyInt = 17
		Const adUnsignedSmallInt = 18
		Const adUnsignedInt = 19
		Const adUnsignedBigInt = 21
		Const adSingle = 4
		Const adDouble = 5
		Const adCurrency = 6
		Const adDecimal = 14
		Const adNumeric = 131
		Const adBoolean = 11
		Const adError = 10
		Const adUserDefined = 132
		Const adVariant = 12
		Const adIDispatch = 9
		Const adIUnknown = 13
		Const adGUID = 72
		Const adDate = 7
		Const adDBDate = 133
		Const adDBTime = 134
		Const adDBTimeStamp = 135
		Const adBSTR = 8
		Const adChar = 129
		Const adVarChar = 200
		Const adLongVarChar = 201
		Const adWChar = 130
		Const adVarWChar = 202
		Const adLongVarWChar = 203
		Const adBinary = 128
		Const adVarBinary = 204
		Const adLongVarBinary = 205
		
		Const adFldIsNullable = &H00000020
		adLockOptimistic = 3
		adOpenKeySet = 3
		adCursorLocation = 3
		
		set dwzRec = server.CreateObject("adodb.recordset")
		dwzRec.LockType = adLockOptimistic
    	dwzRec.CursorType = adOpenKeySet
		dwzRec.CursorLocation = adCursorLocation
		For Each objField In recordset.Fields
			select case clng(objField.Type)
			case 0
				dwzRec.Fields.Append objField.Name, adEmpty, 255, adFldIsNullable
			case 16
				dwzRec.Fields.Append objField.Name, adTinyInt, 255, adFldIsNullable
			case 2
				dwzRec.Fields.Append objField.Name, adSmallInt, 255, adFldIsNullable
			case 3
				dwzRec.Fields.Append objField.Name, adInteger, 255, adFldIsNullable
			case 20
				dwzRec.Fields.Append objField.Name, adBigInt, 255, adFldIsNullable
			case 17
				dwzRec.Fields.Append objField.Name, adUnsignedTinyInt, 255, adFldIsNullable
			case 18
				dwzRec.Fields.Append objField.Name, adUnsignedSmallInt, 255, adFldIsNullable
			case 19
				dwzRec.Fields.Append objField.Name, adUnsignedInt, 255, adFldIsNullable
			case 21
				dwzRec.Fields.Append objField.Name, adUnsignedBigInt, 255, adFldIsNullable
			case 4
				dwzRec.Fields.Append objField.Name, adSingle, 255, adFldIsNullable
			case 5
				dwzRec.Fields.Append objField.Name, adDouble, 255, adFldIsNullable
			case 6
				dwzRec.Fields.Append objField.Name, adCurrency, 255, adFldIsNullable
			case 14
				dwzRec.Fields.Append objField.Name, adDecimal, 255, adFldIsNullable
			case 131
				dwzRec.Fields.Append objField.Name, adNumeric, 255, adFldIsNullable
			case 11
				dwzRec.Fields.Append objField.Name, adBoolean, 255, adFldIsNullable
			case 10
				dwzRec.Fields.Append objField.Name, adError, 255, adFldIsNullable
			case 132
				dwzRec.Fields.Append objField.Name, adUserDefined, 255, adFldIsNullable
			case 12
				dwzRec.Fields.Append objField.Name, adVariant, 255, adFldIsNullable
			case 9
				dwzRec.Fields.Append objField.Name, adIDispatch, 255, adFldIsNullable
			case 13
				dwzRec.Fields.Append objField.Name, adIUnknown, 255, adFldIsNullable
			case 72
				dwzRec.Fields.Append objField.Name, adGUID, 255, adFldIsNullable
			case 7
				dwzRec.Fields.Append objField.Name, adDate, 255, adFldIsNullable
			case 133
				dwzRec.Fields.Append objField.Name, adDBDate, 255, adFldIsNullable
			case 134
				dwzRec.Fields.Append objField.Name, adDBTime, 255, adFldIsNullable
			case 135
				dwzRec.Fields.Append objField.Name, adDBTimeStamp, 255, adFldIsNullable
			case 8
				dwzRec.Fields.Append objField.Name, adBSTR, 255, adFldIsNullable
			case 129
				dwzRec.Fields.Append objField.Name, adChar, 255, adFldIsNullable
			case 200
				dwzRec.Fields.Append objField.Name, adVarChar, 255, adFldIsNullable
			case 201
				dwzRec.Fields.Append objField.Name, adLongVarChar, 255, adFldIsNullable
			case 130
				dwzRec.Fields.Append objField.Name, adWChar, 255, adFldIsNullable
			case 202
				dwzRec.Fields.Append objField.Name, adVarWChar, 255, adFldIsNullable
			case 203
				dwzRec.Fields.Append objField.Name, adLongVarWChar, 255, adFldIsNullable
			case 128
				dwzRec.Fields.Append objField.Name, adBinary, 255, adFldIsNullable
			case 204
				dwzRec.Fields.Append objField.Name, adVarBinary, 255, adFldIsNullable
			case 205
				dwzRec.Fields.Append objField.Name, adLongVarBinary, 255, adFldIsNullable
			case else
				dwzRec.Fields.Append objField.Name, adVarChar, 255, adFldIsNullable
			end select
		next
		if PagingType = "L" or PagingType = "S" then
			dwzRec.Fields.Append "dwzFilter", adVarChar, 255, adFldIsNullable
		end if
		
		dwzRec.open
		
		
		
		while not recordset.eof
			dwzRec.addNew
			For Each objField In recordset.Fields
				dwzRec.Fields.item(objField.Name).value = recordset.Fields.item(objField.Name).value
			next
			if PagingType = "S" then
				tmp = trim(Ucase(cstr(recordset.Fields.item(RecField).value & "")))
				if tmp <> "" then
					dwzRec.Fields.item("dwzFilter").value = lcase(tmp)
				else
					dwzRec.Fields.item("dwzFilter").value = ""
				end if
				if not StringExist(tmp) then
					nFields = nFields + 1
					Redim Preserve Fields(nFields)
					Fields(nFields) = tmp
				end if
			elseif PagingType = "L" then
				tmp = trim(Ucase(cstr(recordset.Fields.item(RecField).value & "")))
				if tmp <> "" then
					dwzRec.Fields.item("dwzFilter").value = lcase(left(tmp, 1))
				else
					dwzRec.Fields.item("dwzFilter").value = ""
				end if
				
				select case left(tmp,1)
				case "A"
					index = 0
				case "B"
					index = 1
				case "C"
					index = 2
				case "D"
					index = 3
				case "E"
					index = 4
				case "F"
					index = 5
				case "G"
					index = 6
				case "H"
					index = 7
				case "I"
					index = 8
				case "J"
					index = 9
				case "K"
					index = 10
				case "L"
					index = 11
				case "M"
					index = 12
				case "N"
					index = 13
				case "O"
					index = 14
				case "P"
					index = 15
				case "Q"
					index = 16
				case "R"
					index = 17
				case "S"
					index = 18
				case "T"
					index = 19
				case "U"
					index = 20
				case "V"
					index = 21
				case "W"
					index = 22
				case "X"
					index = 23
				case "Y"
					index = 24
				case "Z"	
					index = 25
				case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
					index = 26
				case else
					index = 27
				end select
				Letters(index) = Letters(index) + 1
			end if
			dwzRec.update
			recordset.MoveNext
		wend
		
		if PagingType = "S" then
			if CurrentField > nFields then
				CurrentField = 0
			end if
			Filtro = "dwzFilter = '" & Fields(CurrentField) & "'"
			dwzRec.Filter = Filtro
			
		elseif PagingType = "L" then			
			

			
			dwzRec.Filter = Filtro
		else
			maxRecord = dwzRec.RecordCount
			dwzRec.PageSize = RecPaging
			dwzRec.AbsolutePage = CurrentPage
			TotalPage = int(dwzRec.RecordCount / dwzRec.PageSize)
			if (dwzRec.RecordCount / dwzRec.PageSize) - int(dwzRec.RecordCount / dwzRec.PageSize) > 0 then
				TotalPage = TotalPage + 1
			end if	
		end if		
		'response.Write(maxRecord & "<br>" & TotalPage & "<br>" & CurrentPage)
		'response.End()
		
	End Sub
	
	Private function StringExist(str)
		retStr = false
		n = -1
		err.clear
		on error resume next
		n = ubound(Fields)		
		on error goto 0
		if err.number<>0 then
			n = -1
		end if
		for J=0 to n
			if Fields(J) = str then
				retStr = true
				exit for
			end if
		next
		StringExist = retStr
	end function
	
	public function getRecordset()
		set getRecordset = dwzRec
	end function
	
end class

%>
