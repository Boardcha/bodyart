<%
'**********************************
' http://www.dwzone.it
' Csv Import
' Copyright (c) DwZone.it 2000-2005
'**********************************
class dwzCsvImport
	
	Private CsvContent
	Private debug
	Private StartOnEvent
	Private StartOnValue
	Private RedirectPage
	Private DisplayErrors
	Private FieldSeparator
	Private SkipFirstLine
	Private FilePath
	Private FullFilePath
	Private DbFormatDate
	Private DbFormatBoolean
	
	Private ConnectionString
	Private Table
	Private TableUniqueKey
	Private ColIsNum
	Private OnDuplicateEntry
	Private CsvUniqueKey
	Private CsvData()
	
	Private StringNull
	Private StringDelimChar
	Private NumberNull
	Private DateDelimChar
	Private DateNull
	Private EncloseField
	
	Private ItemFieldRec()
	Private ItemFormat()
	Private ItemFromValue()
	Private ItemReference()
	Private ItemColumn()
	
	Private ItemCount
	
	Private TotalRows
	Private ErrMsg
	
	Private Conn
	
	Public sub addItem(Rec, From, Format, Reference, Column)
		ItemCount = ItemCount + 1
		ReDim Preserve ItemFieldRec(ItemCount)
		ReDim Preserve ItemFromValue(ItemCount)
		ReDim Preserve ItemFormat(ItemCount)
		ReDim Preserve ItemReference(ItemCount)
		ReDim Preserve ItemColumn(ItemCount)
		
		ItemFieldRec(ItemCount) = Rec
		ItemFromValue(ItemCount) = From
		ItemFormat(ItemCount) = Format
		ItemReference(ItemCount) = Reference
		ItemColumn(ItemCount) = Column
	end sub
	
	public sub SetEncloseField(param)
		if Ucase(param) = "SA" then
			EncloseField = "'"
		elseif Ucase(param) = "DA" then
			EncloseField = chr(34)
		end if
	end sub
	
	public sub SetStartOn(param1,param2)
		StartOnEvent = param1
		StartOnValue = param2
	end sub
	
	Public sub SetRedirectPage(param)
		RedirectPage = param
	end sub
	
	Public sub SetDisplayErrors(param)
		if lcase(param)="true" then
			DisplayErrors = true
		else
			DisplayErrors = false
		end if
	end sub
	
	Public sub SetFilePath(param)
		FilePath = param
	end sub
		
	public sub SetConnection(param)
		ConnectionString = param
	end sub
	
	public sub SetTable(param)
		Table = param
	end sub
	
	public sub SetTableUniqueKey(param)
		TableUniqueKey = param
	end sub
	
	public sub SetColIsNum(param)
		if lcase(param)="true" then
			ColIsNum = true
		else
			ColIsNum = false
		end if
	end sub
	
	public sub SetOnDuplicateEntry(param)
		OnDuplicateEntry = param
	end sub
	
	public sub SetCsvUniqueKey(param)
		CsvUniqueKey = param
	end sub
	
	public sub SetDbFormatDate(param)
		DbFormatDate = param
	end sub
	
	public sub SetDbFormatBoolean(param)
		DbFormatBoolean = param
	end sub	
	
	public sub SetStringNull(param)
		StringNull = param
	end sub
	
	public sub SetStringDelimChar(param)
		StringDelimChar = param
	end sub
	
	public sub SetNumberNull(param)
		NumberNull = param
	end sub
	
	public sub SetDateDelimChar(param)
		DateDelimChar = param
	end sub
	
	public sub SetDateNull(param)
		DateNull = param
	end sub
	
	Public sub SetFieldSeparator(param)
		if lcase(param) = "tab" then
			FieldSeparator = vbtab
		else
			FieldSeparator = param
		end if
	end sub	
	
	Public sub SetSkipFirstLine(param)
		if lcase(param) = "true" then
			SkipFirstLine = true
		else
			SkipFirstLine = false
		end if
	end sub
	
	Private Sub Class_Initialize()
		
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	
	Public Sub Init ()
		debug = true
		ItemCount = -1
		TotalRows = -1
		StringNull = "Null"
		StringDelimChar = "'"
		NumberNull = "Null"
		DateDelimChar = "'"
		DateNull = "Null"
		set ErrMsg = new XmlError
	End Sub	
	
	private function Start()
		retStr = false
		select case Ucase(StartOnEvent)
		case "GET"
			if request.QueryString(StartOnValue)<>"" then
				retStr = true
			end if
		case "POST"
			if request.Form(StartOnValue)<>"" then
				retStr = true
			end if
		case "REQUEST"
			if request(StartOnValue)<>"" then
				retStr = true
			end if
		case "SESSION"
			if session(StartOnValue)<>"" then
				retStr = true
			end if
		case "COOKIE"
			if request.Cookies(StartOnValue)<>"" then
				retStr = true
			end if
		case "APPLICATION"
			if Application(StartOnValue)<>"" then
				retStr = true
			end if
		case "ONLOAD"
			retStr = true
		end select
		Start = retStr
	end function
	
	
	Public sub Execute()
		if not Start then
			exit sub
		end if
		if ItemCount < 0 then
			Response.write "<strong>DwZone - CSV Import Error.</strong><br/>No Item defined"
			Exit sub
		end if
		if FilePath = "" then
			Response.write "<strong>DwZone - CSV Import Error.</strong><br/>The file path is void"
			Exit sub
		end if
		
		FullFilePath = getFilePath()
		
		if ErrMsg.hasErrors() then
			ResponseError()
			exit sub
		end if
		
		readFromCsvContent()
		
		if ErrMsg.hasErrors() then
			ResponseError()
			exit sub
		end if
		
		set Conn = server.CreateObject("adodb.connection")
		Conn.open ConnectionString
		
		ImportData()
		
		if ErrMsg.hasErrors() then
			ResponseError()
			exit sub
		end if
		
		if RedirectPage <> "" then
			response.Redirect(RedirectPage)
			response.End()
		end if
		
	end sub
	
	private sub ResponseError()
		if DisplayErrors then
			response.write(ErrMsg.getTable("Some errors in the CsvImport"))	
		else
			if RedirectPage <> "" then
				response.Redirect(RedirectPage)
				response.End()
			end if
		end if
	end sub
	
	private sub ImportData()
		for J=0 to TotalRows
			StrSQL = ""
			if OnDuplicateEntry = "NoVerify" then
				StrSQL = CreateInsertQuery(J)
			else
				recExists = RecordExists(J)
				if recExists then
					if OnDuplicateEntry = "Update" then
						StrSQL = CreateUpdateQuery(J)
					elseif OnDuplicateEntry = "Skip" then
						StrSQL = ""
					else
						ErrMsg.add "110", "Row: " & J & " - Err: Duplicate entry", "ImportData"
					end if
				else
					StrSQL = CreateInsertQuery(J)
				end if
			end if
			if StrSQL <> "" then
				err.clear
				on error resume next
				Conn.execute StrSQL
				if err.number<>0 then
					Msg = "Row: " & J & " - Err: " & err.description & "<br>" & StrSQL
					ErrMsg.add Err.number,Msg,"ImportData"
				end if
				on error goto 0
			end if
		next
	end sub
	
	private function CreateUpdateQuery(index)
		sql = "UPDATE " & Table & " set "
		cong = ""
		for K=0 to ItemCount
			if lcase(ItemFieldRec(K)) <> lcase(TableUniqueKey) then	
				strValue = getCsvData(index,ItemReference(K),ItemFromValue(K),ItemColumn(K))
				strValue = FormatValue(strValue, ItemFormat(K))
				sql = sql & cong & ItemFieldRec(K) & "=" & strValue
				cong = ", "
			end if
		next
		sql = sql & " where " & TableUniqueKey
		if ColIsNum then
			Sql = Sql & " = " & getCsvData(index,"","csv",CsvUniqueKey)
		else
			Sql = Sql & " = '" & getCsvData(index,"","csv",CsvUniqueKey) & "'"
		end if
		'response.write Sql
		'response.End()
		CreateUpdateQuery = sql
	end function
	
	
	private function CreateInsertQuery(index)
		sql = "INSERT INTO " & Table & " ("
		cong = ""
		for K=0 to ItemCount
			sql = sql & cong & ItemFieldRec(K)
			cong = ","
		next
		sql = sql & ") VALUES ("
		cong = ""
		for K=0 to ItemCount
			strValue = getCsvData(index,ItemReference(K),ItemFromValue(K),ItemColumn(K))
			strValue = FormatValue(strValue, ItemFormat(K))
			sql = sql & cong & strValue
			cong = ","
		next
		sql = sql & ")"
		'response.write sql
		'response.End()
		CreateInsertQuery = sql
	end function
	
	private function RecordExists(index)
		set Rs = server.CreateObject("adodb.recordset")
		Sql = "select " & TableUniqueKey
		Sql = Sql & " from " & Table
		Sql = Sql & " where " & TableUniqueKey
		if ColIsNum then
			Sql = Sql & " = " & getCsvData(index,"","csv",CsvUniqueKey)
		else
			Sql = Sql & " = '" & getCsvData(index,"","csv",CsvUniqueKey) & "'"
		end if
		'response.write Sql
		'response.End()
		
		Rs.Open Sql,Conn
		if Rs.eof then
			RecordExists = false
		else
			RecordExists = true
		end if
		Rs.close
		set Rs = nothing
	end function
	
	private function getCsvData(index,strValue,ValueFrom,Column)
		'ValueFrom = ""	
		if ValueFrom = "" then
			for kk=0 to ItemCount
				if lcase(ItemReference(kk)) = lcase(strValue) then
					ValueFrom = lcase(ItemFromValue(kk))
					exit for
				end if
			next
		end if
		'response.write ItemReference(K)
		'response.End()
		
		retStr = ""
		select case lcase(ValueFrom)
		case "csv"
			on error resume next
			retStr = CsvData(index)(cstr(clng(Column)-1))
			on error goto 0
		case "get"
			retStr = request.QueryString(ItemReference(K))
		case "post"
			retStr = request.Form(ItemReference(K))
		case "request"
			retStr = request(ItemReference(K))
		case "application"
			retStr = Application(ItemReference(K))
		case "session"
			retStr = Session(ItemReference(K))
		case "cookie"
			retStr = request.Cookies(ItemReference(K))
		case "entered"
			retStr = ItemReference(K)
		end select
		getCsvData = retStr
	end function
	
	private sub readFromCsvContent()
		set Fs = server.CreateObject("Scripting.FileSystemObject")
		set myFile = Fs.openTextFile(FullFilePath, 1)
		LineNumber = 0
		
		if SkipFirstLine then
			tmp = myFile.ReadLine
			LineNumber = LineNumber + 1
		end if
		
		'if EncloseField <> "" then
		'	FieldSeparator = EncloseField & FieldSeparator & EncloseField
		'end if
		
		while not myFile.AtEndOfStream
			LineNumber = LineNumber + 1
			on error resume next
			lineText = trim(myFile.ReadLine)
			if lineText <> "" then
				TotalRows = TotalRows + 1
				ReDim Preserve CsvData(TotalRows)
				set CsvData(TotalRows) = Server.CreateObject("Scripting.Dictionary")
				if EncloseField = "" then
					value = split(lineText, FieldSeparator)
				else
					value = mySplit(lineText, FieldSeparator, EncloseField)
				end if
				for J=0 to ubound(value)
					nName = cstr(J)
					'nValue = value(J),J,ubound(value))
					CsvData(TotalRows)(nName) = value(J)
				next
			end if
			if Err.Number <> 0 then
				ErrMsg.add Err.number,Err.Description,"readFromCsvContent - Line " & LineNumber 
			end if
			on error goto 0
		wend	
	end sub
	
	private function mySplit(str, sSep, sEnc)
		nextChar = ""
		startPos = 1
		uscire = false
		n = -1
		dim retStr()
		
		'1_1;'a___a';'2___2';3_3;'b___b';'Name_1';6___1;'a@aaa.com';6__5;'q__q';'USA'

		do while not uscire
			if nextChar = "" then
				if mid(str, startPos, 1) = sEnc then
					nextChar = sEnc
					startPos = startPos + 1
				else
					nextChar = sSep
				end if
			else
				if mid(str, startPos, 1) = sEnc then
					nextChar = sEnc
					startPos = startPos + 1
				else
					nextChar = sSep
				end if				
			end if
			posNext = instr(startPos, str, nextChar, vbtextcompare)
			
			if posNext < 1 then
				posNext = len(str) + 1
			end if

			n = n + 1
			Redim Preserve retStr(n)
			retStr(n) = mid(str, startPos, posNext - startPos)
						
			if nextChar = sEnc then
				startPos = posNext + 2
			else
				startPos = posNext + 1
			end if
					
			'if n=99 then
			'	response.write str & "<br>" & retStr(9) & "<br>" & startPos & "___" & nextChar & "___" & posNext & "<br>" & startPos & "<br>" & retStr(n) & "<br>" & len(str)
			'	response.End()
			'end if
			
			if startPos >= len(str) then
				if right(str,1) = sSep then
					n = n + 1
					Redim Preserve retStr(n)
					retStr(n) = ""
				end if
				uscire = true
			end if
			
			'if n>=19 then
			'	exit do
			'end if
		loop
		'for J=0 to ubound(retStr)
		'	response.write("-" & retStr(J) & "<br>")
		'next
		'response.End()
		
		mySplit = retStr
	end function
	
	private function RemoveEnclose(val,current,last)
		if EncloseField <> "" then
			if current = 0 then
				if left(val,1) = EncloseField then
					val = mid(val,2)
				end if
			elseif current = last then
				if right(val,1) = EncloseField then
					val = left(val,len(val)-1)
				end if
			end if
		end if
		RemoveEnclose = val
	end function
	
	private function getFilePath()
		if FilePath = "" then
			ErrMsg.add "100","The file is missing","getFilePath"
			getFilePath = ""
			exit function
		else
			set Fs = server.CreateObject("Scripting.FileSystemObject")
			if not Fs.FileExists(server.MapPath(FilePath)) then
				ErrMsg.add "101","The file: " & server.MapPath(FilePath) & " is not find","getFilePath"
				getFilePath = ""
				exit function	
			end if
			set Fs = nothing
			getFilePath = server.MapPath(FilePath)
		end if
	end function
	
	Private Function escapeValue(valueStr)
		If valueStr & "" = "" Then 
			escapeValue = ""
		else
			valueStr = valueStr & ""
			valueStr = replace(valueStr, "'", "''")
			escapeValue = valueStr
		End If		
	End Function

	
	
	private function FormatValue(strValue, Format)
		if lcase(left(Format,1)) = lcase("S") then
			FormatValue = FormatAsString(strValue, Format)
			
		elseif lcase(left(Format,1)) = lcase("N") then
			FormatValue = FormatAsNumber(strValue, Format)
			
		elseif lcase(left(Format,1)) = lcase("D") then
			FormatValue = FormatAsDate(strValue, Format)
			
		elseif lcase(left(Format,1)) = lcase("B") then
			FormatValue = FormatAsChk(strValue, Format)
			
		end if		
	end function
	
	private function FormatAsChk(strValue,format)
		'B,-1,0
		tmp = split(format,",")
		tmpDbValue = split(dbFormatBoolean,"/")
		strValue = strValue & ""
		if lcase(strValue) = lcase(tmp(1)) then
			FormatAsChk = tmpDbValue(0)
		else
			FormatAsChk = tmpDbValue(1)
		end if
	end function
	
	private function FormatAsDate(strValue,format)
		'D,',none,NULL
		strValue = strValue & ""
		if strValue <> "" then
			if isDate(strValue) then
				myDate = cdate(strValue)
				strValue = replace(DbFormatDate,"DD",right("0" & day(myDate),2),1,-1,vbtextcompare)
				strValue = replace(strValue,"MM",right("0" & month(myDate),2),1,-1,vbtextcompare)
				strValue = replace(strValue,"YYYY",year(myDate),1,-1,vbtextcompare)
				FormatAsDate = DateDelimChar & strValue & DateDelimChar
			else
				FormatAsDate = DateNull
			end if
		else
			FormatAsDate = DateNull
		end if
	end function
	
	private function FormatAsNumber(strValue,format)
		'N,none,none,NULL
		strValue = strValue & ""
		if strValue<>"" then
			if format = "N." then
				strValue = replace(strValue,",","")
			elseif format = "N," then
				strValue = replace(replace(strValue,".",""),",",".")
			end if
			if isnumeric(strValue) then
				FormatAsNumber = strValue
			else
				FormatAsNumber = NumberNull
			end if
		else
			FormatAsNumber = NumberNull
		end if
	end function
	
	private function FormatAsString(strValue,format)
		strValue = strValue & ""
		if strValue = "" then
			FormatAsString = StringNull
		else
			FormatAsString = StringDelimChar & escapeValue(strValue) & StringDelimChar
		end if
	end function
	
end class

%>
<!--#include file="dwzExportUtils.asp"-->