<%
'--------------------------------------------------------------------------------------------
' Caller Service
' ==============
' This page will be included in all pages.It will give service of calling the sdk api using
' credential details provided in the Constants file.
'--------------------------------------------------------------------------------------------
	Dim gv_APIEndpoint
	Dim gv_nvpHeader
	Dim gv_ProxyServer	
	Dim gv_ProxyServerPort 
	Dim gv_Proxy
	
      gv_APIEndpoint	= API_ENDPOINT
	
      'WinObjHttp Request proxy settings.
	gv_ProxyServer	= HTTPREQUEST_PROXYSETTING_SERVER
	gv_ProxyServerPort = HTTPREQUEST_PROXYSETTING_PORT
	gv_Proxy		= 2	'setting for proxy activation
	gv_UseProxy		= USE_PROXY
     

'----------------------------------------------------------------------------------
' Purpose: Make the API call to PayPal, using API signature.
' Inputs:  Method name to be called & NVP string to be sent with the post method
' Returns: NVP Collection object of Call Response.
'----------------------------------------------------------------------------------	
Function hash_call ( methodName,nvpStr )
On Error Resume Next
	Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")

      nvpStr	=	Server.URLEncode("METHOD")&"=" &Server.URLEncode(methodName)&nvpStr  

      Set SESSION("nvpReqArray")= deformatNVP(nvpStr)
	objHttp.open "POST", gv_APIEndpoint, False
	WinHttpRequestOption_SslErrorIgnoreFlags=4
	objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
	
	
	If  gv_UseProxy Then
	'Proxy Call
	objHttp.SetProxy gv_Proxy,  gv_ProxyServer& ":" &gv_ProxyServerPort
	End If
	
	objHttp.Send nvpStr
	
	Set SESSION("nvpReqArray")= deformatNVP(nvpStr)
	Set  nvpResponseCollection =deformatNVP(objHttp.responseText)
	Set  hash_call =nvpResponseCollection
	Set objHttp = Nothing 
	
	If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"hash_call")
	SESSION("nvpReqArray") =  Null
'	Response.Redirect "http://" &  Request.ServerVariables("server_name")  & "/PayPal/APIError.asp"
	Else
	SESSION("ErrorMessage")	= Null
	End If
	
	
End Function
'----------------------------------------------------------------------------------
' Purpose: It will Formates error Messages.
' Inputs:  NVP string.
' Returns: NVP Collection object deformated from NVP string.
'----------------------------------------------------------------------------------
Function ErrorFormatter ( errDesc,errNumber,errSource,errlocation )
	ErrorFormatter ="Error Description : " &errDesc& "<br/>Error number : " &errNumber& "<br/>Error Source : " &errSource& "<br/>Error Location : " &errlocation& ""
End Function 
'----------------------------------------------------------------------------------
' Purpose: It will convert nvp string to Collection object.
' Inputs:  NVP string.
' Returns: NVP Collection object deformated from NVP string.
'----------------------------------------------------------------------------------
Function deformatNVP ( nvpstr )
On Error Resume Next
	Dim AndSplitedArray,EqualtoSplitedArray,Index1,Index2,NextIndex
	Set NvpCollection = Server.CreateObject("Scripting.Dictionary")
	AndSplitedArray = Split(nvpstr, "&", -1, 1)
	NextIndex=0
	For Index1 = 0 To UBound(AndSplitedArray)
		EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
		For Index2 = 0 To UBound(EqualtoSplitedArray)
			NextIndex=Index2+1
			NvpCollection.Add URLDecode(EqualtoSplitedArray(Index2)),URLDecode(EqualtoSplitedArray(NextIndex))
			Index2=Index2+1
		Next
	Next
	Set deformatNVP = NvpCollection
	If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"deformatNVP")
'	Response.Redirect "http://" &  Request.ServerVariables("server_name")  & "/PayPal/APIError.asp"
	else
	SESSION("ErrorMessage")	= Null
	End If
End Function
 '----------------------------------------------------------------------------------
' Purpose: It gives out encoded url path to dispaly.
' Inputs:  Url string to be encoded.
' Returns: Decoded Url string.
'----------------------------------------------------------------------------------
Function URLEncode(str) 
On Error Resume Next
    Dim AndSplitedArray,EqualtoSplitedArray,Index1,Index2,UrlEncodeString,NvpUrlEncodeString
	AndSplitedArray = Split(nvpstr, "&", -1, 1)
	UrlEncodeString=""
	NvpUrlEncodeString=""
	For Index1 = 0 To UBound(AndSplitedArray)
		EqualtoSplitedArray=Split(AndSplitedArray(Index1), "=", -1, 1)
		For Index2 = 0 To UBound(EqualtoSplitedArray)
		If Index2 = 0 then
			UrlEncodeString=UrlEncodeString & Server.URLEncode(EqualtoSplitedArray(Index2))
		Else			
			UrlEncodeString=UrlEncodeString &"="& Server.URLEncode(EqualtoSplitedArray(Index2))
		End if
		Next
		If Index1 = 0 then
			NvpUrlEncodeString= NvpUrlEncodeString & UrlEncodeString
		Else			
			NvpUrlEncodeString= NvpUrlEncodeString &"&"&UrlEncodeString
		End if
		UrlEncodeString=""
	Next
	URLEncode = NvpUrlEncodeString
	
	If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLEncode")
'	Response.Redirect "http://" &  Request.ServerVariables("server_name")  & "/PayPal/APIError.asp"
	else
	SESSION("ErrorMessage")	= Null
	End If
	
 End Function 
'----------------------------------------------------------------------------------
' Purpose: It gives out decoded url path to dispaly.
' Inputs:  Url string to be decoded.
' Returns: Decoded Url string.
'----------------------------------------------------------------------------------
 Function URLDecode(byVal encodedstring)
	Dim strIn, strOut, intPos, strLeft
	Dim strRight, intLoop
	strIn  = encodedstring : strOut = _
		 "" : intPos = Instr(strIn, "+")
	Do While intPos
		strLeft = "" : strRight = ""
		If intPos > 1 then _
			strLeft = Left(strIn, intPos - 1)
		If intPos < len(strIn) then _
			strRight = Mid(strIn, intPos + 1)
		strIn = strLeft & " " & strRight
		intPos = InStr(strIn, "+")
		intLoop = intLoop + 1
	Loop
	intPos = InStr(strIn, "%")
	Do while intPos
		If intPos > 1 then _
			strOut = strOut & _
				Left(strIn, intPos - 1)
		strOut = strOut & _
			Chr(CInt("&H" & _
				mid(strIn, intPos + 1, 2)))
		If intPos > (len(strIn) - 3) then
			strIn = ""
		Else
			strIn = Mid(strIn, intPos + 3)
		End If
		intPos = InStr(strIn, "%")
	Loop
	URLDecode = strOut & strIn
 	If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLDecode")
'	Response.Redirect "http://" &  Request.ServerVariables("server_name")  & "/PayPal/APIError.asp"
	else
	SESSION("ErrorMessage")	= Null
	End If
End Function

'----------------------------------------------------------------------------------
' Purpose: It gives url path for the cancel & return  page.
' Returns: Url path of current page without file name.
'----------------------------------------------------------------------------------
Function GetURL() 
On Error Resume Next
server_name= Request.ServerVariables("SERVER_NAME")
Path= Request.ServerVariables("PATH_INFO")
Port_number= request.servervariables("SERVER_PORT") 
If (Request.ServerVariables("HTTPS") = "off") Then
    htt="http://"
else
    htt="https://"
End If
Full_path= htt & server_name &":"&Port_number&Path
urlParts = split(Full_path,"/") 
Virtual_Path = trim(Replace(Full_path,urlParts(Ubound(urlParts)),"" ))
GetURL = Virtual_Path
	If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"hash_call")
'	Response.Redirect "http://" &  Request.ServerVariables("server_name")  & "/PayPal/APIError.asp"
	else
	SESSION("ErrorMessage")	= Null
	End If
End Function
'----------------------------------------------------------------------------------
' Purpose: It's Workaround Method for Response.Redirect
'          It will redirect the page to the specified url without urlencoding
' Inputs: Url to redirect the page
'----------------------------------------------------------------------------------
Function ReDirectURL(url)

On Error Resume Next

response.clear
response.status="302 Object moved"
response.AddHeader "location",url
response.flush
	If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"ReDirectURL")
'	Response.Redirect "http://" &  Request.ServerVariables("server_name")  & "/PayPal/APIError.asp"
	else
	SESSION("ErrorMessage")	= Null
	End If
End Function

%>
