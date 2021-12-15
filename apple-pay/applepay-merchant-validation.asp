<!--#include virtual="/connections/chilkat.asp"-->
<%

Mode = "TEST" ' OR PRODUCTION
'Note: Apple Pay does not work on localhost since they require domain verification.

If Mode = "PRODUCTION" Then domain = "bodyartforms.com" else domain = "dev5.bodyartforms.com"

set objUrl = Server.CreateObject("Chilkat_9_5_0.Url")
validation_url = Request.QueryString("url")
success = objUrl.ParseUrl(validation_url)

If Right(objUrl.Host, 9) = "apple.com" Then
	options = "{" & _
				"merchantIdentifier: ""merchant.com.bodyartforms""," & _
				"displayName: ""Bodyartforms""," & _
				"initiative: ""web""," & _
				"initiativeContext: """ & domain & """" & _
			"}"

	url = "https://apple-pay-gateway.apple.com/paymentservices/paymentSession"

	set cert = Server.CreateObject("Chilkat_9_5_0.Cert")

	'====== To load certificate from Windows ======
	'Certificate must be imported to Windows
	'success = cert.LoadByCommonName("Apple Pay Merchant Identity:merchant.com.bodyartforms")
	'If (success <> 1) Then
	'    Response.Write "<pre>" & Server.HTMLEncode( cert.LastErrorText) & "</pre>"
	'    Response.End
	'End If

	'If (cert.HasPrivateKey() <> 1) Then
	'    Response.Write "<pre>" & Server.HTMLEncode( "A private key is needed for TLS client authentication.") & "</pre>"
	'    Response.Write "<pre>" & Server.HTMLEncode( "This certificate has no private key.") & "</pre>"
	'    Response.End
	'End If
	'=====================================================================
	
	'====== To load certificate from a .pfx file ======
	'THIS FILE MUST BE MOVED TO OUT OF PUBLIC FOLDERS
	pfxFilename = Server.MapPath("merchant_validation_cert.pfx")

	' A PFX typically contains certificates in the chain of authentication.
	' The Chilkat cert object will choose the certificate w/
	' private key farthest from the root authority cert.
	' To access all the certificates in a PFX, use the 
	' Chilkat certificate store object instead.
	success = cert.LoadPfxFile(pfxFilename, "ugurcatak")
	If (success <> 1) Then
		Response.Write "<pre>" & Server.HTMLEncode( cert.LastErrorText) & "</pre>"
		Response.End
	End If

	set http = Server.CreateObject("Chilkat_9_5_0.Http")

	http.SetRequestHeader "Content-Type", "application/json"
	http.Accept = "application/json"

	success = http.SetSslClientCert(cert)

	Set resp = http.PostJson2(url,"application/json", options)
	If (http.LastMethodSuccess = 0) Then
		Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
		Response.End
	End If

	jsonResponseStr = resp.BodyStr
	Response.Write jsonResponseStr
End If
%>