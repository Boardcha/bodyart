<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/admin/etsy-v3/etsy-refresh-token.asp" -->
<%
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

autoReconnect = 1
tls = 1
success = rest.Connect("openapi.etsy.com",443,tls,autoReconnect)
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
success = sbAuthHeaderVal.Append("Bearer ")
success = sbAuthHeaderVal.Append(etsy_access_token)
rest.Authorization = sbAuthHeaderVal.GetAsString() 

success = rest.AddQueryParam("client_id", etsy_consumer_key)
success = rest.AddQueryParam("was_shipped",0)
success = rest.AddQueryParam("limit",100)

jsonResponseText = rest.FullRequestNoBody("GET","/v3/application/shops/" & etsy_baf_shop_id & "/receipts")
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

set jsonResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonResponse.Load(jsonResponseText)
jsonResponse.EmitCompact = 0

'Response.Write "<pre>" & Server.HTMLEncode( jsonResponse.Emit()) & "</pre>"
'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

i = 0
count_i = jsonResponse.SizeOfArray("results")

Do While i < count_i
	jsonResponse.I = i
		
	var_receipt_id = jsonResponse.StringOf("results[i].receipt_id") 

	'========== RETRIEVE ORDER FROM DATABASE IF TRACK IS AVAILABLE ==================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT shipped, transactionID, shipping_type, USPS_tracking, pay_method, date_sent, email, customer_first, customer_last from sent_items WHERE USPS_tracking <> '' AND transactionID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
	set rsGetEtsyTracks = objCmd.Execute()

	IF NOT rsGetEtsyTracks.EOF THEN

		if InStr(rsGetEtsyTracks.Fields.Item("shipping_type").Value, "DHL") > 0 then
			carrier_name = "dhl-global-mail"
		end if
		if InStr(rsGetEtsyTracks.Fields.Item("shipping_type").Value, "USPS") > 0 then
			carrier_name = "usps"
		end if


		'======= CLEAR PARAMS FROM PRIOR SHIPMENT =============================
		success = rest.ClearAllQueryParams()

		success = rest.AddQueryParam("client_id", etsy_consumer_key)
		success = rest.AddQueryParam("tracking_code",rsGetEtsyTracks.Fields.Item("USPS_tracking").Value)
		success = rest.AddQueryParam("carrier_name",carrier_name)

		jsonResponseText = rest.FullRequestNoBody("POST","/v3/application/shops/" & etsy_baf_shop_id & "/receipts/" & rsGetEtsyTracks.Fields.Item("transactionID").Value & "/tracking")
		If (rest.LastMethodSuccess = 0) Then
			Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
			Response.End
		End If

		response.write "<br>" & jsonResponseText

	END IF '==== if order has been shipped in database and has track #
	set rsGetEtsyTracks = nothing
	i = i + 1
Loop

Set rest = Nothing
Set sbAuthHeaderVal = Nothing
Set jsonResponse = Nothing
Set objCmd = Nothing
DataConn.Close

%>
