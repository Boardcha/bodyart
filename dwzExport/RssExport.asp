<%
'**********************************
' http://www.dwzone.it
' Rss Writer
' Copyright (c) DwZone.it 2000-2005
'**********************************
class dwzRssExport
	
	Private Encoding
	Private CodePage
	Private PubDate
	Private Title
	Private Description
	Private Link
	Private Recordset
	Private ItemTitle
	Private ItemDescription 
	Private ItemLink
	Private ItemLinkText
	Private ItemAuthor
	Private ItemPubDate
	Private FileName
	Private NumberOfRecord
	Private StartOnEvent
	Private StartOnValue
	Private FeedImageUrl
	Private FeedImageLink
	Private FeedImageTitle
	Private Language
	Private Category
	Private TTL
	Private Docs
	Private ManagingEditor
	Private Webmaster
	Private Generator
	Private TimeZone
	
	Private ItemLabelField()
	Private ItemRecField()
	Private ItemTypeTag()
	Private ItemAttrName()
	Private ItemAttrValue()
	Private ItemCustomCount
	
	Private BaseLabelField()
	Private BaseRecField()
	Private BaseTypeTag()
	Private BaseAttrName()
	Private BaseAttrValue()
	Private BaseCustomCount
	
	
	
	Public sub addCustomItem(Label, Rec, sType, sAttrName, sAttrValue)
		ItemCustomCount = ItemCustomCount + 1
		ReDim Preserve ItemLabelField(ItemCustomCount)
		ReDim Preserve ItemRecField(ItemCustomCount)
		ReDim Preserve ItemTypeTag(ItemCustomCount)
		ReDim Preserve ItemAttrName(ItemCustomCount)
		ReDim Preserve ItemAttrValue(ItemCustomCount)
		
		ItemLabelField(ItemCustomCount) = Label
		ItemRecField(ItemCustomCount) = Rec
		ItemTypeTag(ItemCustomCount) = sType
		ItemAttrName(ItemCustomCount) = sAttrName
		ItemAttrValue(ItemCustomCount) = sAttrValue
		
	end sub
	
	public sub addBaseCustomItem(Label, Rec, sType, sAttrName, sAttrValue)
		BaseCustomCount = BaseCustomCount + 1
		ReDim Preserve BaseLabelField(BaseCustomCount)
		ReDim Preserve BaseRecField(BaseCustomCount)
		ReDim Preserve BaseTypeTag(BaseCustomCount)
		ReDim Preserve BaseAttrName(BaseCustomCount)
		ReDim Preserve BaseAttrValue(BaseCustomCount)
		
		BaseLabelField(BaseCustomCount) = Label
		BaseRecField(BaseCustomCount) = Rec
		BaseTypeTag(BaseCustomCount) = sType
		BaseAttrName(BaseCustomCount) = sAttrName
		BaseAttrValue(BaseCustomCount) = sAttrValue
	end sub
	
	Public sub setRecordset (ByRef rs)
		Set recordset = rs
	End Sub
	
	Public Sub setEncoding(param)
		tmp = split(param,";")
		Encoding = tmp(0)
		CodePage = tmp(1)
	End Sub	
	
	Public Sub setPubDate(param)
		if lcase(param) = "true" then
			PubDate = true
		else
			PubDate = false
		end if
	End Sub	
	
	Public Sub SetTitle(param)
		Title = getEvalValue(param)
	End Sub	
	
	Public Sub SetDescription(param)
		Description = getEvalValue(param)
	End Sub	
	
	Public Sub SetLink(param)
		Link = getEvalValue(param)
	End Sub	
	
	Public Sub SetItemTitle(param)
		ItemTitle = Trim(param)
	End Sub	
	
	Public Sub SetItemDescription(param)
		ItemDescription = Trim(param)
	End Sub	
	
	Public Sub SetItemLink(param)
		ItemLink = Trim(param)
	End Sub	
	
	Public sub SetItemLinkText(param)
		ItemLinkText = trim(param)
	end sub
	
	Public Sub SetItemAuthor(param)
		ItemAuthor = Trim(param)
	End Sub	
	
	Public Sub SetItemPubDate(param)
		ItemPubDate = Trim(param)
	End Sub	
	
	Public Sub SetFileName(param)
		FileName = Trim(param)
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
	
	public sub SetFeedImage(param)
		tmp = split(param,"__@_@__")
		FeedImageUrl = getEvalValue(tmp(0))
		FeedImageLink = getEvalValue(tmp(1))
		FeedImageTitle = getEvalValue(tmp(1))
	end sub
	
	public sub SetAdditionalInfo(param)
		tmp = split(param,"__@_@__")
		Language = tmp(0)
		Category = getEvalValue(tmp(1))
		TTL = getEvalValue(tmp(2))
		Docs = getEvalValue(tmp(3))
		ManagingEditor = getEvalValue(tmp(4))
		Webmaster = getEvalValue(tmp(5))
		Generator = getEvalValue(tmp(6))
	end sub
	
	Public sub SetTimeZone(param)
		TimeZone = param
	end sub
	
	Private Sub Class_Initialize()
		
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	
	Public Sub Init ()
		Set Recordset = nothing	
		ItemCustomCount = -1
		BaseCustomCount = -1
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
			Response.write "<strong>DwZone - RSS Export Error.</strong><br/>The recordset is not valid"
			Exit sub
		End If
		
		Set cont = new XmlContent
		cont.Add("<" & "?xml version=""" & "1.0" & """ encoding=""" & Encoding & """?>" & vbcrlf)
		cont.Add("<rss version=""" & "2.0" & """>" & vbcrlf)
		cont.Add(vbTab & "<channel>" & vbcrlf)
		if Title<>"" then
			cont.Add(vbTab & "<title>" & escapeValue(Title) & "</title>" & vbcrlf)
		end if
		if link<>"" then
			cont.Add(vbTab & "<link>" & escapeValue(link) & "</link>" & vbcrlf)
		end if
		if description<>"" then
			cont.Add(vbTab & "<description>" & escapeValue(description) & "</description>" & vbcrlf)
		end if
		if language<>"" then
			cont.Add(vbTab & "<language>" & escapeValue(language) & "</language>" & vbcrlf)
		end if
		if PubDate then
			'Mon, 28 Feb 2005 13:12:09 +0100
			cont.Add(vbTab & "<pubDate>" & getStringDate(Now) & "</pubDate>" & vbcrlf)
		end if
		if docs<>"" then
			cont.Add(vbTab & "<docs>" & escapeValue(docs) & "</docs>" & vbcrlf)
		end if
		if managingEditor<>"" then
			cont.Add(vbTab & "<managingEditor>" & escapeValue(managingEditor) & "</managingEditor>" & vbcrlf)
		end if
		if webMaster<>"" then
			cont.Add(vbTab & "<webMaster>" & escapeValue(webMaster) & "</webMaster>" & vbcrlf)
		end if
		if category<>"" then
			cont.Add(vbTab & "<category>" & escapeValue(category) & "</category>" & vbcrlf)
		end if
		if TTL<>"" then
			cont.Add(vbTab & "<ttl>" & escapeValue(TTL) & "</ttl>" & vbcrlf)
		end if
		if generator<>"" then
			cont.Add(vbTab & "<generator>" & escapeValue(generator) & "</generator>" & vbcrlf)
		end if	
		
		for J=0 to BaseCustomCount
			if BaseTypeTag(J) = "Date" then
				tmpVal = getEvalValue(BaseRecField(J))
				if isDate(tmpVal) then
					tmpVal = getStringDate(cdate(tmpVal))
				else
					tmpVal = ""
				end if
			elseif BaseTypeTag(J) = "Text" then
				tmpVal = escapeValue(getEvalValue(BaseRecField(J)))
			else
				tmpVal = getEvalValue(BaseRecField(J))
			end if
			if BaseAttrName(J) <> "" then
					Attribute = " " & BaseAttrName(J) & "=" & chr(34) & searchValue(BaseAttrValue(J)) & chr(34) & " "
				else
					Attribute = ""
				end if
			if BaseTypeTag(J) = "Memo" then
				cont.Add(vbTab & "<" & BaseLabelField(J) & Attribute & "><![CDATA[" & tmpVal & "]]></" & BaseLabelField(J) & ">" & vbcrlf)
			else
				cont.Add(vbTab & "<" & BaseLabelField(J) & Attribute & ">" & tmpVal & "</" & BaseLabelField(J) & ">" & vbcrlf)
			end if
		next
		
		if FeedImageUrl<>"" then
			cont.Add(vbTab & "<image>" & vbcrlf)
			cont.Add(vbTab & vbTab & "<title>" & FeedImageTitle & "</title>" & vbcrlf)
			cont.Add(vbTab & vbTab & "<url>" & getFullLinK(FeedImageUrl) & "</url>" & vbcrlf)	
			cont.Add(vbTab & vbTab & "<link>" & FeedImageLink & "</link>" & vbcrlf)					
			cont.Add(vbTab & "</image>" & vbcrlf)
		end if
		
		nRec = 0
		do while not Recordset.eof
			cont.Add(vbTab & vbTab & "<item>" & vbcrlf)
			if ItemTitle<>"" then
				cont.Add(vbTab & vbTab & vbTab & "<title>" & escapeValue(Recordset(ItemTitle)) & "</title>" & vbcrlf)
			end if
			if ItemLink<>"" or ItemLinkText<>"" then
				if ItemLink<>"" then
					cont.Add(vbTab & vbTab & vbTab & "<link>" & ItemLinkText & escapeValue(Recordset(ItemLink)) & "</link>" & vbcrlf)
				else
					cont.Add(vbTab & vbTab & vbTab & "<link>" & ItemLinkText & "</link>" & vbcrlf)
				end if
			end if
			if ItemDescription<>"" then
				cont.Add(vbTab & vbTab & vbTab & "<description><![CDATA[" & Recordset(ItemDescription) & "]]></description>" & vbcrlf)
			end if
			if ItemAuthor<>"" then
				cont.Add(vbTab & vbTab & vbTab & "<author>" & escapeValue(Recordset(ItemAuthor)) & "</author>" & vbcrlf)
			end if
			if ItemPubDate<>"" then
				if isDate(Recordset(ItemPubDate)) then
					cont.Add(vbTab & vbTab & vbTab & "<pubDate>" & getStringDate(cdate(Recordset(ItemPubDate))) & "</pubDate>" & vbcrlf)
				else
					cont.Add(vbTab & vbTab & vbTab & "<pubDate></pubDate>")
				end if
			end if
			
			for J=0 to ItemCustomCount				
				if ItemTypeTag(J) = "Date" then
					tmpVal = Recordset(ItemRecField(J))
					if isDate(tmpVal) then
						tmpVal = getStringDate(cdate(tmpVal))
					else
						tmpVal = ""
					end if
				elseif ItemTypeTag(J) = "Text" then
					tmpVal = escapeValue(Recordset(ItemRecField(J)))
				else
					tmpVal = Recordset(ItemRecField(J))
				end if
				if ItemAttrName(J) <> "" then
					Attribute = " " & ItemAttrName(J) & "=" & chr(34) & searchValue(ItemAttrValue(J)) & chr(34) & " "
				else
					Attribute = ""
				end if
				if ItemTypeTag(J) = "Memo" then
					cont.Add(vbTab & vbTab & vbTab & "<" & ItemLabelField(J) & Attribute & "><![CDATA[" & tmpVal & "]]></" & ItemLabelField(J) & ">" & vbcrlf)
				else
					cont.Add(vbTab & vbTab & vbTab & "<" & ItemLabelField(J) & Attribute & ">" & tmpVal & "</" & ItemLabelField(J) & ">" & vbcrlf)
				end if
			next
			
			cont.Add(vbTab & vbTab & "</item>" & vbcrlf)
			nRec = nRec + 1
			if Ucase(NumberOfRecord) <> "ALL" then
				if clng(NumberOfRecord) = nRec then
					exit do
				end if
			end if
			Recordset.MoveNext
		loop
		
		cont.Add(vbTab & "</channel>" & vbcrlf)
		cont.Add("</rss>" & vbcrlf)
		
		RssContent = cont.ToString
		
		Response.Clear()
		Response.AddHeader "Pragma", "public"
		Response.AddHeader "Expires", "Thu, 19 Nov 1981 08:52:00 GMT"
		Response.AddHeader "Cache-Control", "must-revalidate, post-check=0, pre-check=0"
		Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
		Response.AddHeader "Cache-Control", "private"
		'response.codepage = CodePage
		Response.CharSet = Encoding
		Response.ContentType = "text/xml"
		Response.AddHeader "Content-Length", len(RssContent)
		If FileName <> "" Then
			Response.AddHeader "Content-disposition", "attachment; filename=""" & FileName & """;"
		End If
		Response.write RssContent
		Response.Flush()
		Response.End()		
		
	end sub
	
	private function searchValue(str)
		retStr = ""
		err.clear
		on error resume next
		retStr = Recordset(str)
		if err.number <> 0 then
			retStr = str
		end if
		on error goto 0
		searchValue = retStr
	end function
	
	private function getStringDate(mydate)
		tmp = ""
		tmp = getLetterDay(WeekDay(mydate,vbsunday))
		tmp = tmp & ", " & Day(mydate)
		tmp = tmp & " " & getLetterMonth(Month(mydate))
		tmp = tmp & " " & Year(mydate)
		tmp = tmp & " " & right("0" & Hour(mydate),2)
		tmp = tmp & ":" & right("0" & Minute(mydate),2)
		tmp = tmp & ":" & right("0" & second(mydate),2)
		tmp = tmp & " " & TimeZone
		getStringDate = tmp
	end function
	
	private function getLetterDay(d)
		tmpDay = split(",Sun,Mon,Tue,Wed,Thu,Fri,Sat",",")	
		getLetterDay = tmpDay(clng(d))
	end function
	
	private function getLetterMonth(m)
		tmpMonth = split(",Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",")		
		getLetterMonth = tmpMonth(clng(m))
	end function
		
	Private Function escapeValue(valueStr)
		If valueStr & "" = "" Then 
			escapeValue = ""
		else
			valueStr = valueStr & ""
			valueStr = replace(valueStr, chr(34), "''")
			'valueStr = replace(valueStr, "&", "&amp;")
			valueStr = replace(valueStr, ">", "")
			valueStr = replace(valueStr, "<", "")
			valueStr = replace(valueStr, vbcrlf, "")
			valueStr = replace(valueStr, vbcr, "")
			valueStr = replace(valueStr, vblf, "")
		End If	
		escapeValue = valueStr
	End Function
	
	private function getEvalValue(valueStr)
		retStr = ""
		If isNull(valueStr) Or len(valueStr)=0 or valueStr="" Then 
			retStr = ""
		else
			retStr = replace(valueStr,"@_''_@",chr(34))
			while instr(retStr,"@_start_@")>0
				inizio = instr(retStr,"@_start_@") + 9
				fine = instr(inizio,retStr,"@_end_@")
				lung = fine-inizio
				retStr = left(retStr,inizio-10) & eval(mid(retStr,inizio,lung)) & mid(retStr,fine + 7)
			wend
		end if
		getEvalValue = retStr
	end function
	
	private function getFullLink(LinkVal)
		if lcase(left(LinkVal,7))<>"http://" then
			if right(LinkVal,1)<>"/" then
				LinkVal = LinkVal & "/"
			end if
			LinkVal = "http://" & Request.ServerVariables("SERVER_NAME") & LinkVal
		end if
		getFullLink = LinkVal
	end function
	
end class

%>
<!--#include file="dwzExportUtils.asp"-->