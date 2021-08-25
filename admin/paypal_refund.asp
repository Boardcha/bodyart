<%@LANGUAGE="VBSCRIPT" CodePage=65001%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!-- #include file ="../Connections/paypal-constants.asp.asp" -->
<!-- #include file ="../PayPal/CallerService.asp" -->
<!-- #include file ="../PayPal/DisplayAllResponse.asp" -->
<%

Response.Charset = "UTF-8"

	Dim transactionID
	Dim refundType
	Dim currencyCode
	Dim note
	Dim amount
      Dim gv_APIUserName
	Dim gv_APIPassword
	Dim gv_APISignature
	Dim gv_Version
      Dim gv_SUBJECT


	transactionID		= Request.Form ("transactionID")
	refundType			= Request.Form ("refundType")
	currencyCode		= "USD"
	note				= Request.Form ("memo")
	amount			= Request.Form ("amount")
      gv_APIUserName	      = API_USERNAME
	gv_APIPassword	      = API_PASSWORD
	gv_APISignature         = API_SIGNATURE
	gv_Version		      = API_VERSION
	gv_SUBJECT              = SUBJECT
      

'-----------------------------------------------------------------------------
' Construct the request string that will be sent to PayPal.
' The variable $nvpstr contains all the variables and is a
' name value pair string with &as a delimiter
'-----------------------------------------------------------------------------

	If refundType="Partial" Then	

	nvpstr	=	"&REFUNDTYPE="&refundType &_
				"&NOTE="&note &_
				"&CURRENCYCODE="&currencyCode &_
				"&AMT=" &amount &_
				"&TRANSACTIONID=" &transactionID 
	Else
		nvpstr	=	"&TRANSACTIONID=" &transactionID & _
				"&REFUNDTYPE="&refundType &_
				"&NOTE="&note &_
				"&CURRENCYCODE="&currencyCode 
	End If
				
	nvpstr	=	URLEncode(nvpstr)

     If IsEmpty(gv_SUBJECT) Then
      
     nvpStr =nvpstr&"&USER=" & gv_APIUserName &_
                              "&PWD=" &gv_APIPassword &_
                              "&SIGNATURE=" & gv_APISignature &_
                              "&VERSION=" & gv_Version

     ElseIf IsEmpty(gv_APIUserName )and IsEmpty(gv_APIPassword) and IsEmpty(gv_APISignature) Then

     nvpStr =nvpstr&"&SUBJECT=" & gv_SUBJECT &_
                              "&VERSION=" & gv_Version

     Else
     
     nvpStr =nvpstr&"&USER=" & gv_APIUserName &_
                              "&PWD=" &gv_APIPassword &_
                              "&SIGNATURE=" & gv_APISignature &_
                              "&VERSION=" & gv_Version &_
                              "&SUBJECT=" & gv_SUBJECT 
     End If
	
	
'-----------------------------------------------------------------------------
' Make the API call to PayPal,using API signature.
' The API response is stored in an associative array called gv_resArray
'-----------------------------------------------------------------------------
	Set resArray	= hash_call("RefundTransaction",nvpstr)
	ack = UCase(resArray("ACK"))
	amt = UCase(resArray("GROSSREFUNDAMT"))

%>
<html>
<link href="../includes/nav.css" rel="stylesheet" type="text/css">
<title>PayPal refund</title>
<body class="mainbkgd">
<%
'----------------------------------------------------------------------------------
' Display the API request and API response back to the browser.
' If the response from PayPal was a success, display the response parameters
' If the response was an error, display the errors received
'----------------------------------------------------------------------------------
	If ack="SUCCESS" Then
		Message ="Transaction refunded!!"
%>
<% 
if request.form("Type") = "Backorder" then

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

commUpdate.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0, item_price = 0, notes = 'Out; Refunded $" & Request.Form("x_amount") & " " & date () & "' WHERE OrderDetailID = " & Request.Form("OrderDetailID") 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()
Set commUpdate = Nothing

end if
%>
</p>
<p class="adminheader"><a href="invoice.asp?ID=<%= request.form("invoice") %>">CLICK HERE</a> to view updated invoice</p>
<%	Else ' if transaction failed
		 Set SESSION("nvpErrorResArray") = resArray
		 Response.Redirect "http://www.bodyartforms.com/PayPal/APIError.asp"
	End If

%>
</body>
</html>

