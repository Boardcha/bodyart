<%@LANGUAGE="VBSCRIPT" CodePage = 65001 %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/endicia.asp" -->
<%
'IIS should process this page as 65001 (UTF-8), responses should be 
'treated as 28591 (ISO-8859-1).
Response.CharSet = "ISO-8859-1"
Response.CodePage = 28591
%>
<%

XML =	"<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
					   "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""" & _
					   "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		  "<soap:Body>" & _
			"<GetSCAN xmlns=""www.envmgr.com/LabelService"">" & _
			  "<GetSCANRequest>" & _
				"<RequesterID>" & endicia_requester_id & "</RequesterID>" & _
				"<RequestID>" & RandomID(50) & "</RequestID>" & _
				"<CertifiedIntermediary>" & _
				  "<AccountID>" & endicia_account_id & "</AccountID>" & _
				  "<PassPhrase>" & endicia_passphrase & "</PassPhrase>" & _ 
				"</CertifiedIntermediary>" & _
				"<GetSCANRequestParameters ImageResolution=""96"" ImageFormat=""PDF"">" & _
						  "<FromName>Body Art Forms</FromName>" & _
						  "<FromCompany>Body Art Forms</FromCompany>" & _				  
						  "<FromAddress>1966 S Austin Ave</FromAddress>" & _
						  "<FromCity>Georgetown</FromCity>" & _
						  "<FromState>TX</FromState>" & _
						  "<FromZip>78626</FromZip>" & _
				"</GetSCANRequestParameters>" & _
				"<ManifestType>USPS</ManifestType>" & _
				"<NumberOfContainerLabels>1</NumberOfContainerLabels>" & _
			  "</GetSCANRequest>" & _
			"</GetSCAN>" & _   
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
	http.SetRequestHeader "SOAPAction","www.envmgr.com/LabelService/GetSCAN"
	
	
	' Some services expect the content-type in the HTTP header to be "application/xml" while
	' other expect text/xml.  The default sent by Chilkat is "application/xml", 
	' but Endicia web service expects "text/xml".  Therefore, change the content-type:
	http.SetRequestHeader "Content-Type","text/xml; charset=utf-8"

	' The endpoint for this soap service is:
	endPoint = endicia_url

	' resp is a Chilkat_9_5_0.HttpResponse
	Set resp = http.PostXml(endPoint,strXml,"utf-8")
	If (http.LastMethodSuccess <> 1) Then
		httpError = httpError & "<br><pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If

	set xmlResp = Server.CreateObject("Chilkat_9_5_0.Xml")
	success = xmlResp.LoadXml(resp.BodyStr)
	'Response.Write "<pre>" & Server.HTMLEncode( xmlResp.GetXml()) & "</pre>"
	
	responseStatusCode = resp.StatusCode
	if responseStatusCode = 200 Then
		base64PDF = xmlResp.AccumulateTagContent("SCANForm","")
		SubmissionID = xmlResp.AccumulateTagContent("SubmissionID","")
		If base64PDF <> "" Then
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
			objCmd.CommandText = "INSERT INTO tbl_closeout_forms (provider, requestId, usps_base64_manifest_pdf) VALUES ('USPS', ?, ?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("requestId",200,1,1000, SubmissionID))
			objCmd.Parameters.Append(objCmd.CreateParameter("usps_base64_manifest_pdf",200,1,-1, base64PDF))
			objCmd.Execute()
			
			json_status = "success"
			json_message = "success"
		Else
			json_status = "error"
			If httpError<>"" Then json_message = json_message & httpError
			json_message = json_message & xmlResp.AccumulateTagContent("ErrorMessage","")
		End If
	Else
		json_status = "error"
		If httpError<>"" Then json_message = json_message & httpError
		json_message = json_message & xmlResp.AccumulateTagContent("ErrorMessage","")
	End If
%>
<%
Function RandomID(size)
    Const VALID_TEXT = "abcdefghijklmnopqrstuvwxyz1234567890"
    Dim Length, sNewSearchTag, I

    Length = Len(VALID_TEXT)

    Randomize()

    For I = 1 To size            
        sNewSearchTag = sNewSearchTag & Mid(VALID_TEXT, Int(Rnd()*Length + 1), 1)
    Next

    RandomID = sNewSearchTag
End Function
%>
{
    "status":"<%=json_status%>",
    "message": "<%=json_message%>"
}
