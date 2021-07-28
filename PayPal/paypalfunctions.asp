<%
	' ===================================================
	' PayPal API Include file
	' 
	' Defines all the global variables and the wrapper functions 
	'-----------------------------------------------------------

	Dim gv_APIEndpoint
	Dim gv_APIUser
	Dim gv_APIPassword
	Dim gv_APIVendor
	Dim gv_APIPartner
	Dim gv_BNCode

	Dim gv_ProxyServer	
	Dim gv_ProxyServerPort 
	Dim gv_Proxy		
	
	'----------------------------------------------------------------------------------
	' Authentication Credentials for making the call to the server
	'----------------------------------------------------------------------------------
	SandboxFlag = true
	Env = "live" ' pilot or live
	
	'------------------------------------
	' PayPal API Credentials 
	'------------------------------------
	gv_APIUser	= "bodyartforms"
	gv_APIPassword	= "DRAjufU8athED4"
	gv_APIVendor = "bodyartforms"
	gv_APIPartner = "PayPal"

	'----------------------------------------------------------------------
	' Define the PayPal URLs for Express Checkout 
	' 	This is the URL that the buyer is first sent to do authorize payment with their paypal account
	' 	change the URL depending if you are testing on the sandbox
	' 	or going to the live PayPal site
	'
	' For the sandbox, the URL is       https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token=
	' For the live site, the URL is     https://www.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token=
	'------------------------------------------------------------------------
	if Env = "pilot" Then
		gv_APIEndpoint = "https://pilot-payflowpro.paypal.com"
		PAYPAL_URL = "https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token="
	Else
		gv_APIEndpoint = "https://payflowpro.paypal.com"
		PAYPAL_URL = "https://www.paypal.com/cgi-bin/webscr?cmd=_express-checkout&token="
	End If 
		
	'WinObjHttp Request proxy settings.
	gv_ProxyServer	= "127.0.0.1"
	gv_ProxyServerPort = "808"
	gv_Proxy		= 2	'setting for proxy activation
	gv_UseProxy		= False
	
	

	'-------------------------------------------------------------------------------------------------------------------------------------------
	' Purpose: 	Prepares the parameters for direct payment (credit card) and makes the call.
	'
	' Note:
	'		There are other optional inputs for credit card processing that are not presented here.
	'		For a complete list of inputs available, please see the documentation here for US and UK:
	'		https://cms.paypal.com/cms_content/US/en_US/files/developer/PP_PayflowPro_Guide.pdf
	'		https://cms.paypal.com/cms_content/GB/en_GB/files/developer/PP_WebsitePaymentsPro_IntegrationGuide_UK.pdf
	'		
	' Returns: 
	'		The NVP Collection object of the Response.
	'--------------------------------------------------------------------------------------------------------------------------------------------	
	
	Function DirectPayment( paymentType, paymentAmount, creditCardNumber, expDate, cvv2, firstName, lastName, street, city, billstate, zip, countryCode, currencyCode, orderdescription, billingphone, shiptofirstname, shiptolastname, shiptoaddress, shiptocity, shiptostate, shiptozip, shiptocountry, invoicenum, billtoemail, trxtype, trans_id )

		nvpstr = "&TENDER=C"  ' C is for credit card tender type
		nvpstr = nvpstr & "&TRXTYPE=" & trxtype ' S stands for Sale, A stands for Authorization, C = credit

		'unique request ID
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		unique_id = TypeLib.Guid

		nvpstr = nvpstr & "&ACCT=" & creditCardNumber
		nvpstr = nvpstr & "&CVV2=" & cvv2
		nvpstr = nvpstr & "&EXPDATE=" & expDate
	'	nvpstr = nvpstr & "&ACCTTYPE=" & creditCardType
		nvpstr = nvpstr & "&AMT=" & paymentAmount
		nvpstr = nvpstr & "&CURRENCY=USD"
		nvpstr = nvpstr & "&BILLTOFIRSTNAME=" & firstName
		nvpstr = nvpstr & "&BILLTOLASTNAME=" & lastName
		nvpstr = nvpstr & "&BILLTOSTREET=" & street
		nvpstr = nvpstr & "&BILLTOCITY=" & city
		nvpstr = nvpstr & "&BILLTOSTATE=" & billstate
		nvpstr = nvpstr & "&BILLTOZIP=" & zip
		nvpstr = nvpstr & "&BILLTOCOUNTRY=" & countryCode
		nvpstr = nvpstr & "&SHIPTOSTATE=" & shiptostate
		nvpstr = nvpstr & "&SHIPTOZIP=" & shiptozip
		nvpstr = nvpstr & "&SHIPTOCOUNTRY=" & shiptocountry
		nvpstr = nvpstr & "&SHIPTOFIRSTNAME=" & shiptofirstName
		nvpstr = nvpstr & "&SHIPTOLASTNAME=" & shiptolastName
		nvpstr = nvpstr & "&SHIPTOSTREET=" & shiptostreet
		nvpstr = nvpstr & "&SHIPTOCITY=" & shiptocity
		nvpstr = nvpstr & "&BILLTOEMAIL=" & billtoemail
		nvpstr = nvpstr & "&BILLTOPHONENUM=" & billingphone
		nvpstr = nvpstr & "&INVNUM=" & invoicenum
		nvpstr = nvpstr & "&ORDERDESC=" & orderdescription
		nvpstr = nvpstr & "&ORIGID=" & trans_id	
		
		
		' Transaction results (especially values for declines and error conditions) returned by each PayPal-supported
		' processor vary in detail level and in format. The Payflow Verbosity parameter enables you to control the kind
		' and level of information you want returned. 
		' By default, Verbosity is set to LOW. A LOW setting causes PayPal to normalize the transaction result values. 
		' Normalizing the values limits them to a standardized set of values and simplifies the process of integrating 
		' the Payflow SDK.
		' By setting Verbosity to MEDIUM, you can view the processor’s raw response values. This setting is more “verbose”
		' than the LOW setting in that it returns more detailed, processor-specific information. 
		' Review the chapter in the Developer's Guides regarding VERBOSITY and the INQUIRY function for more details.
		' Set the transaction verbosity to MEDIUM.
		nvpstr = nvpstr & "&VERBOSITY=HIGH"

		'-------------------------------------------------------------------------------------------
		' Make the call to Payflow to finalize payment
		' If an error occured, show the resulting errors
		'-------------------------------------------------------------------------------------------
		set DirectPayment = hash_call(nvpstr,unique_id)
	End Function
	
	
	'----------------------------------------------------------------------------------
	' Purpose: 	Make the API call to PayPal, using API signature.
	' Inputs:  
	'		Method name to be called & NVP string to be sent with the post method
	' Returns: 
	'		NVP Collection object of Call Response.
	'----------------------------------------------------------------------------------	
	Function hash_call ( nvpStr, unique_id )
		Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")

		nvpStrComplete = "USER=" & gv_APIUser & "&VENDOR=" & gv_APIVendor & "&PARTNER=" & gv_APIPartner & "&PWD=" & gv_APIPassword & nvpStr 
		nvpStrComplete	= nvpStrComplete & "&BUTTONSOURCE=" & Server.URLEncode( gv_BNCode )
		
		Set SESSION("nvpReqArray")= deformatNVP( nvpStrComplete )
		objHttp.open "POST", gv_APIEndpoint, False
		WinHttpRequestOption_SslErrorIgnoreFlags=4
		objHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
		
		objHttp.SetRequestHeader "Content-Type", "text/namevalue"
		objHttp.SetRequestHeader "Content-Length", Len(nvpStrComplete)
		objHttp.SetRequestHeader "X-VPS-CLIENT-TIMEOUT", "45" 
		objHttp.SetRequestHeader "X-VPS-REQUEST-ID", unique_id 
		' set the host header
		If Env = "pilot" Then
			objHttp.SetRequestHeader "Host", "pilot-payflowpro.paypal.com"
		Else
			objHttp.SetRequestHeader "Host", "payflowpro.paypal.com"
		End If
		
		If gv_UseProxy Then
			'Proxy Call
			objHttp.SetProxy gv_Proxy,  gv_ProxyServer& ":" &gv_ProxyServerPort
		End If
		
		objHttp.Send nvpStrComplete
				
		Set nvpResponseCollection = deformatNVP(objHttp.responseText)
		Set hash_call = nvpResponseCollection
		Set objHttp = Nothing 
		
		If Err.Number <> 0 Then 
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"hash_call")
			SESSION("nvpReqArray") =  Null
		Else
			SESSION("Message")	= Null
		End If
	End Function

	'----------------------------------------------------------------------------------
	' Purpose: 	Formats the error Messages.
	' Inputs:  
	'		
	' Returns: 
	'		Formatted Error string
	'----------------------------------------------------------------------------------
	Function ErrorFormatter ( errDesc, errNumber, errSource, errlocation )
		ErrorFormatter ="<font color=red>" & _
								"<TABLE align = left>" &_
								"<TR>" &"<u>Error Occured!!!</u>" & "</TR>" &_
								"<TR>" &"<TD>Error Description :</TD>" &"<TD>"&errDesc& "</TD>"& "</TR>" &_
								"<TR>" &"<TD>Error number :</TD>" &"<TD>"&errNumber& "</TD>"& "</TR>" &_
								"<TR>" &"<TD>Error Source :</TD>" &"<TD>"&errSource& "</TD>"& "</TR>" &_
								"<TR>" &"<TD>Error Location :</TD>" &"<TD>"&errlocation& "</TD>"& "</TR>" &_
								"</TABLE>" &_
								"</font>"
	End Function 

	'----------------------------------------------------------------------------------
	' Purpose: 	Convert nvp string to Collection object.
	' Inputs:  	
	'		NVP string.
	' Returns: 
	'		NVP Collection object created from deserializing the NVP string.
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
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"deformatNVP")
		else
			SESSION("Message")	= Null
		End If
	End Function

	'----------------------------------------------------------------------------------
	' Purpose: URL Encodes a string
	' Inputs:  
	'		String to be url encoded.
	' Returns: 
	'		Url Encoded string.
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
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLEncode")
		else
			SESSION("Message")	= Null
		End If
		
	 End Function 

	'----------------------------------------------------------------------------------
	' Purpose: Decodes a URL Encoded string
	' Inputs:  
	'		A URL encoded string
	' Returns: 
	'		Decoded string.
	'----------------------------------------------------------------------------------
	Function URLDecode(str) 
		On Error Resume Next
		
		str = Replace(str, "+", " ") 
		For i = 1 To Len(str) 
			sT = Mid(str, i, 1) 
			If sT = "%" Then 
				If i+2 < Len(str) Then 
					sR = sR & _ 
						Chr(CLng("&H" & Mid(str, i+1, 2))) 
					i = i+2 
				End If 
			Else 
				sR = sR & sT 
			End If 
		Next 
	       
		URLDecode = sR 
		If Err.Number <> 0 Then 
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"URLDecode")
		else
			SESSION("Message")	= Null
		End If

	End Function

	'----------------------------------------------------------------------------------
	' Purpose: 	It's Workaround Method for Response.Redirect
	'          	It will redirect the page to the specified url without urlencoding
	' Inputs: 
	'		Url to redirect the page
	'----------------------------------------------------------------------------------
	Function ReDirectURL( token )
		On Error Resume Next

		payPalURL = PAYPAL_URL & token
		response.clear
		response.status="302 Object moved"
		response.AddHeader "location", payPalURL
		If Err.Number <> 0 Then 
			SESSION("Message")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"ReDirectURL")
		else
			SESSION("Message")	= Null
		End If
	End Function	
	
%>