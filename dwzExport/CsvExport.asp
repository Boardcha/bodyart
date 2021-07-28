<%
'**********************************
' http://www.dwzone.it
' Csv Writer
' Copyright (c) DwZone.it 2000-2005
'**********************************
class dwzCsvExport
	
	Private Recordset
	Private FileName
	Private NumberOfRecord
	Private StartOnEvent
	Private StartOnValue
	Private FieldSeparator
	Private FieldLabel
	
	Private ItemLabel()
	Private ItemRecField()
	Private ItemFormat()
	Private ItemCount

	
	Public sub addItem(Label, Rec, Format)
		ItemCount = ItemCount + 1
		ReDim Preserve ItemLabel(ItemCount)
		ReDim Preserve ItemRecField(ItemCount)
		ReDim Preserve ItemFormat(ItemCount)
		
		ItemLabel(ItemCount) = Label
		ItemRecField(ItemCount) = Rec
		ItemFormat(ItemCount) = Format
	end sub
	
	Public sub setRecordset (ByRef rs)
		Set recordset = rs
	End Sub
		
	Public Sub SetFileName(param)
		if Trim(param) <> "" then
			FileName = Trim(param)
		else
			FileName = "Export.csv"
		end if
	End Sub	
	
	Public Sub SetNumberOfRecord(param)
		if Ucase(param) <> "ALL" then
			if not isnumeric(param) then
				param = "ALL"
			end if
		end if
		NumberOfRecord = Trim(param)
	End Sub	
	
	public sub SetStartOn(param1,param2)
		StartOnEvent = param1
		StartOnValue = param2
	end sub
	
	Public sub SetFieldSeparator(param)
		if Ucase(param) = "TAB" then
			FieldSeparator = vbtab
		else
			FieldSeparator = param
		end if
	end sub
	
	Public sub SetFieldLabel(param)
		if lcase(param) = "true" then
			FieldLabel = true
		else
			FieldLabel = false
		end if
	end sub
	
	
	Private Sub Class_Initialize()
		
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	
	Public Sub Init ()
		Set Recordset = nothing	
		ItemCount = -1
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
		If Not isObject(recordset) Then
			Response.write "<strong>DwZone - XML Export Error.</strong><br/>The recordset is not valid"
			Exit sub
		End If
		if ItemCount < 0 then
			Response.write "<strong>DwZone - XML Export Error.</strong><br/>No Item defined"
			Exit sub
		end if	
		
		
		Set cont = new XmlContent
		
		if FieldLabel then
			newLine = ""
			for J=0 to ItemCount
				if newLine = "" then
					newLine = escapeValue(ItemLabel(J))
				else
					newLine = newLine & FieldSeparator & escapeValue(ItemLabel(J))
				end if
			next
			cont.Add(newLine & vbcrlf)
		end if
				
		nRec = 0
		do while not Recordset.eof
			
			newLine = ""
			for J=0 to ItemCount
				strValue = Recordset(ItemRecField(J)) & ""
				if strValue <> "" and ItemFormat(J)<>"String" then
					strValue = FormatValue(strValue,ItemFormat(J))
				end if
				strValue = escapeValue(strValue)
				if newLine = "" then
					newLine = strValue
				else
					newLine = newLine & FieldSeparator & strValue
				end if
			next
			cont.Add(newLine & vbcrlf)	
			
			nRec = nRec + 1
			if Ucase(NumberOfRecord) <> "ALL" then
				if clng(NumberOfRecord) = nRec then
					exit do
				end if
			end if
			Recordset.MoveNext
		loop
		
		RssContent = cont.ToString
		
		Response.Clear()
		Response.AddHeader "Pragma", "public"
		Response.AddHeader "Expires", "Thu, 19 Nov 1981 08:52:00 GMT"
		Response.AddHeader "Cache-Control", "must-revalidate, post-check=0, pre-check=0"
		Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
		Response.AddHeader "Cache-Control", "private"
		Response.ContentType = "text/html"
		Response.AddHeader "Content-Length", len(RssContent)
		Response.AddHeader "Content-disposition", "attachment; filename=""" & FileName & """;"
		Response.write RssContent
		Response.Flush()
		Response.End()		
		
	end sub
		
	Private Function escapeValue(valueStr)
		If valueStr & "" = "" Then 
			escapeValue = ""
		else
			valueStr = valueStr & ""
			valueStr = replace(valueStr, FieldSeparator, "")
			valueStr = replace(valueStr, vbcrlf, "")
			valueStr = replace(valueStr, vbcr, "")
			valueStr = replace(valueStr, vblf, "")
			escapeValue = valueStr
		End If		
	End Function

	
	private function FormatValue(strValue, Format)
		if lcase(left(Format,6)) = lcase("Number") then
			FormatValue = FormatAsNumber(strValue, Format)
		elseif lcase(left(Format,4)) = lcase("Date") then
			FormatValue = FormatAsDate(strValue, Format)
		elseif lcase(left(Format,8)) = lcase("Checkbox") then
			FormatValue = FormatAsChk(strValue, Format)
		end if		
	end function
	
	private function FormatAsNumber(strValue, Format)
		retStr = ""
		select case Format
		case "Number (Default)"
			retStr = FormatNumber(strValue)
		case "Number (0 decimal)"
			retStr = FormatNumber(strValue,0,0,0,0)
		case "Number (1 decimal)"
			retStr = FormatNumber(strValue,1,0,0,0)
		case "Number (2 decimal)"
			retStr = FormatNumber(strValue,2,0,0,0)
		case "Number (3 decimal)"
			retStr = FormatNumber(strValue,3,0,0,0)
		case "Number (4 decimal)"
			retStr = FormatNumber(strValue,4,0,0,0)
		case else
			retStr = strValue
		end select
		retStr = CDbl(retStr)
		FormatAsNumber = retStr
	end function
	
	private function FormatAsChk(strValue, Format)
		tmpFormat = replace(Format,"Checkbox","",1,-1,vbtextcompare)
		tmpFormat = replace(tmpFormat,"(","",1,-1,vbtextcompare)
		tmpFormat = replace(tmpFormat,")","",1,-1,vbtextcompare)
		tmpValue = split(trim(tmpFormat),"/")
		if strValue then
			FormatAsChk = tmpValue(0)
		else
			FormatAsChk = tmpValue(1)
		end if
	end function
	
	private function FormatAsDate(strValue, Format)
		if not isDate(strValue) then
			FormatAsDate = ""
			exit function
		end if
		myDate = cdate(strValue)
		retStr = trim(replace(Format,"DATE","",1,-1,vbtextcompare))
		retStr = replace(retStr,"DD",day(myDate),1,-1,vbbinarycompare)
		retStr = replace(retStr,"MM",month(myDate),1,-1,vbbinarycompare)
		retStr = replace(retStr,"YYYY",Year(myDate),1,-1,vbbinarycompare)
		retStr = replace(retStr,"h",hour(myDate),1,-1,vbbinarycompare)
		retStr = replace(retStr,"m",minute(myDate),1,-1,vbbinarycompare)
		retStr = replace(retStr,"s",second(myDate),1,-1,vbbinarycompare)
		FormatAsDate = trim(retStr)
	end function
	
end class

%>
<!--#include file="dwzExportUtils.asp"-->