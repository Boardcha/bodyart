<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/sql_connection.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>PayPal error</title>
<link href="../CSS/stylesheet.css" rel="stylesheet" type="text/css" />
<% if Request.Cookies("LayoutView") = "" then %>
<link href="../CSS/inc_NewLeftnav.css" rel="stylesheet" type="text/css" />
<link href="../CSS/inc_NewProducts.css" rel="stylesheet" type="text/css" />
<% end if %>
<% if Request.Cookies("LayoutView") = "Compact" then %>
<link href="../CSS/inc_NewLeftnav.css" rel="stylesheet" type="text/css" />
<link href="../CSS/inc_CompactProducts.css" rel="stylesheet" type="text/css" />
<% end if %>
<% if Request.Cookies("LayoutView") = "Classic" then %>
<link href="../CSS/inc_ClassicLeftnav.css" rel="stylesheet" type="text/css" />
<link href="../CSS/inc_ClassicProducts.css" rel="stylesheet" type="text/css" />
<% end if %>
</head>

<body  onLoad="javascript:document.forms.FRM_redirect.submit();">
<!--#include file="../inc_header_asp.asp" -->
<!--#include file="../inc_LeftNav.asp" -->
<div id="Content">
  
<div class="LargeHeader">Paypal error</div>
<div class="ContentText">

<%

Response.Charset = "UTF-8"
'--------------------------------------------------------------------------------------------
' API Request and Response/Error Output
' =====================================
' This page will be called after getting Response from the server
' or any Error occured during comminication for all APIs,to display Request,Response or Errors.
'--------------------------------------------------------------------------------------------
	Dim resArray
	Dim message
	Dim ResponseHeader
	Dim Sepration
	On Error Resume Next
	message		 =  SESSION("msg")
	Sepration		=":"
	Set resArray = SESSION("nvpErrorResArray")
	
	ResponseHeader="Error Response Details"
	
	
	If Not  SESSION("ErrorMessage")Then
	message = SESSION("ErrorMessage")
	ResponseHeader=""
	Sepration		=""
	End If
	
	
	If Err.Number <> 0 Then
	
	SESSION("nvpReqArray") = Null
	
	Response.flush
	End If
'--------------------------------------------------------------------------------------------
' If there is no Errors Construct the HTML page with a table of variables Loop through the associative array 
' for both the request and response and display the results.
'--------------------------------------------------------------------------------------------
%>
		<link href="sdk.css" rel="stylesheet" type="text/css">
		<!-- #include file ="CallerService.asp" -->
	</HEAD>
	<body alink="#0000ff" vlink="#0000ff">
		<center>
			<table width="700">
				<tr>
					<td colspan="2" class="header" height="16">
						<%=message%>
					</td>
				</tr>
				<tr>
					<td colspan="2" class="header">
						<center>
							<%=ResponseHeader%>
						</center>
					</td>
				</tr>
				<!--displying all Response parameters -->
				<% 
		    reskey = resArray.Keys
		    resitem = resArray.items
			For resindex = 0 To resArray.Count - 1 
     %>
				<tr>
					<td class="field">
            <% =reskey(resindex) %>
						<B>
							<%=Sepration%>
						</B>
					</td>
					<td>
						<% =resitem(resindex) %>
					</td>
				</tr>
				<% next %>
				</TR>
			</table>
		</center>
		<br>
	
<%
DIM strPage
strPage = Request.QueryString("RecurringPage")
%>
<a class="home" href="../contact.asp" target="_blank"><B>Contact Us<B></a>

</div>
 
</div>
<!-- End main content area -->
<!--#include file="../inc_footer.asp" -->
</body>
</html>
