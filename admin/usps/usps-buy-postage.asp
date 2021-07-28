<%@LANGUAGE="VBSCRIPT" CodePage = 65001 %>
<!--#include virtual="/Connections/endicia.asp" -->
<%
'IIS should process this page as 65001 (UTF-8), responses should be 
'treated as 28591 (ISO-8859-1).
Response.CharSet = "ISO-8859-1"
Response.CodePage = 28591
%>
<%
' ============================================================================
' When this page called once, it deposits $500 into the endicia sandbox account.
' ============================================================================

XML =	"<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
					   "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""" & _
					   "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		  "<soap:Body>" & _
			"<BuyPostage xmlns=""www.envmgr.com/LabelService"">" & _
			  "<RecreditRequest>" & _
				"<RequesterID>" & endicia_requester_id & "</RequesterID>" & _
				"<RequestID>3</RequestID>" & _
				"<CertifiedIntermediary>" & _
				  "<AccountID>" & endicia_account_id & "</AccountID>" & _
				  "<PassPhrase>" & endicia_passphrase & "</PassPhrase>" & _ 
				"</CertifiedIntermediary>" & _
				"<RecreditAmount>500</RecreditAmount>" & _
			  "</RecreditRequest>" & _
			"</BuyPostage>" & _
		  "</soap:Body>" & _
		"</soap:Envelope>"		
		
		
	set http = Server.CreateObject("Chilkat_9_5_0.Http")
	set objXml = Server.CreateObject("Chilkat_9_5_0.Xml")
	success = objXml.LoadXml(XML)
	If (success <> 1) Then
		errorOnGettingLabel = errorOnGettingLabel + 1
		xmlError = xmlError & "<br><pre>" & Server.HTMLEncode( objXml.LastErrorText) & "</pre>"
	End If

	strXml = objXml.GetXml()

	' We'll need to add this in the HTTP header:
	' SOAP Action for BUYPostage
	http.SetRequestHeader "SOAPAction","www.envmgr.com/LabelService/BuyPostage"
	
	
	' Some services expect the content-type in the HTTP header to be "application/xml" while
	' other expect text/xml.  The default sent by Chilkat is "application/xml", 
	' but Endicia web service expects "text/xml".  Therefore, change the content-type:
	http.SetRequestHeader "Content-Type","text/xml; charset=utf-8"

	' The endpoint for this soap service is:
	endPoint = endicia_url

	' resp is a Chilkat_9_5_0.HttpResponse
	Set resp = http.PostXml(endPoint,strXml,"utf-8")
	If (http.LastMethodSuccess <> 1) Then
		ResponseError = ResponseError & "<br><pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If

	responseStatusCode = resp.StatusCode
	if responseStatusCode = 200 Then
		Response.Write "SUCCESSFUL<br><pre>" & Server.HTMLEncode( "XML Response:") & "</pre><br>"
	Else
		Response.Write "<pre>" & Server.HTMLEncode( "Response Status Code: " & responseStatusCode) & "</pre>"
	End If
	
	set xmlResp = Server.CreateObject("Chilkat_9_5_0.Xml")
	success = xmlResp.LoadXml(resp.BodyStr)
	Response.Write "<pre>" & Server.HTMLEncode( xmlResp.GetXml()) & "</pre>"	

%>

