<%
Server.ScriptTimeout = 1000
%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<!--#include virtual="/emails/function-send-email.asp"-->
<%
'=== CHECK DELIVERED ORDERS ===
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ID, customer_first, email, estimated_delivery_date, USPS_tracking FROM sent_items WHERE estimated_delivery_date = CONVERT(VARCHAR(10), GETDATE(), 23) AND delivered_email_sent = 0" 
Set rsDelivered = objCmd.Execute()

While Not rsDelivered.EOF
	status = getDeliveryStatus(rsDelivered("USPS_tracking"))
	If status = "ORDER_DELIVERED" Then 
		GetOrderItems(rsDelivered("ID"))
		mailer_type = "ORDER_DELIVERED"
		var_email = "amanda@bodyartforms.com"
		'rsDelivered("email")
		var_first = rsDelivered("customer_first")
		var_invoiceid = rsDelivered("ID")
		var_tracking = "Your tracking # is <strong>" & rsDelivered("USPS_tracking") & "</strong>. If you have an account on our website, you can track your package by going to your order history and pressing the Track Order button. Or, you can track your package by going directly to <a href=""https://bodyartforms.com/dhl-tracker.asp?tracking=" & rsDelivered("USPS_tracking") & """>this link</a>." & mail_order_details				
		%>
		<!--#include virtual="/emails/email_variables.asp"-->
		<%
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET delivered_email_sent = 1 WHERE ID = " & rsDelivered("ID")
		objCmd.Execute()	
	End If
	rsDelivered.MoveNext
Wend
rsDelivered.Close

Response.Write "Successfuly completed." 
Set rsDelivered = Nothing
DataConn.Close()
Set DataConn = Nothing
%>

<%
'=== FUNCTIONS ===
Function getDeliveryStatus(trackingNumber)

	set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
	bTls = 1
	port = 443
	bAutoReconnect = 1
	success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
	success = rest.ClearAllQueryParams()
	success = rest.AddQueryParam("dhlPackageId", trackingNumber)
	success = rest.AddQueryParam("pickup", "5351961")

	set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sbAuthHeaderVal.Append("Bearer ")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
	rest.Authorization = sbAuthHeaderVal.GetAsString()

	ResponseGetTrack = rest.FullRequestNoBody("GET","/tracking/v4/package")
	set JsonTracking = Server.CreateObject("Chilkat_9_5_0.JsonObject")
	JsonTracking.EmitCompact = 0
	JsonTracking.Load(ResponseGetTrack)
	'Response.Write "<pre>" & Server.HTMLEncode( JsonTracking.Emit()) & "</pre>"

	If JsonTracking.IntOf("packages") <> "" Then 

		If JsonTracking.StringOf("packages[0].events") <> "" Then ' ONLY LOOP If THERE ARE EVENTS TO SHOW 
			Set eventsArray = JsonTracking.ArrayOf("packages[0].events")
			eventsSize = eventsArray.Size
			j = eventsSize - 1

			For e = 0 To eventsSize - 1	
				Set eventsObj = eventsArray.ObjectAt(j)
												   
				If eventsObj.StringOf("primaryEventDescription") = "DELIVERED" Then
					var_status = "ORDER_DELIVERED"
					'======== WRITE DELIVERED STATUS TO DB ===================================
					set objCmd = Server.CreateObject("ADODB.command")
					objCmd.ActiveConnection = DataConn
					objCmd.CommandText = "UPDATE sent_items SET date_delivered = ?, packaged_delivered = ? WHERE USPS_tracking = ?"
					objCmd.Parameters.Append(objCmd.CreateParameter("date_delivered",135,1,30, eventsObj.StringOf("date") & " " & eventsObj.StringOf("time")))
					objCmd.Parameters.Append(objCmd.CreateParameter("packaged_delivered",3,1,2, 1))
					objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking",200,1,200,trackingNumber))
																							
					objCmd.Execute()
				End If
				
				If var_status <> "ORDER_DELIVERED" AND eventsObj.IntOf("primaryEventId") = 598 Then
					var_status = "OUT_FOR_DELIVERY"
				End If
				
				j = j - 1
			Next
		Else ' NO EVENTS FOUND
			var_status = "No tracking information available yet"
		End If ' only show if there are events
	Else ' NO PACKAGE FOUND
		var_status = "No tracking information available yet"
	End If
	
	getDeliveryStatus = var_status
	
End Function

Function GetOrderItems(InvoiceID)
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT InvoiceID, ProductID, DetailID, title, ProductDetail1, Gauge, Length, stock_qty, OrderDetailID, email, customer_first, title, qty, ProductDetail1, ProductDetailID, item_price, PreOrder_Desc, picture, free, type, title FROM dbo.QRY_OrderDetails WHERE InvoiceID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,20, InvoiceID))
	Set rsGetInfo = objCmd.Execute()

	'================================================================================================
	' START store details into a dynamic multidimensional array
	reDim array_details_2(12,0)

		array_gauge = ""
		if rsGetInfo("Gauge") <> "" then
			array_gauge = Server.HTMLEncode(rsGetInfo("Gauge"))
		end if
		
		array_length = ""
		if rsGetInfo("Length") <> "" then
			array_length = Server.HTMLEncode(rsGetInfo("Length"))
		end if
		
		array_detail = ""
		if rsGetInfo("ProductDetail1") <> "" then
			array_detail = Server.HTMLEncode(rsGetInfo("ProductDetail1"))
		end if
		
		array_add_new = uBound(array_details_2,2) 
		REDIM PRESERVE array_details_2(12,array_add_new+1) 

		array_details_2(0,array_add_new) = rsGetInfo("ProductDetailID")
		array_details_2(1,array_add_new) = rsGetInfo("qty")
		array_details_2(2,array_add_new) = rsGetInfo("title") 
		array_details_2(3,array_add_new) = array_gauge
		array_details_2(4,array_add_new) = FormatNumber(rsGetInfo("item_price"), -1, -2, -2, -2)
		
		var_preorder_text = ""
		if rsGetInfo("PreOrder_Desc") <> "" then
			var_preorder_text = replace(rsGetInfo("PreOrder_Desc"),"{}", "   ")
		end if
		
		array_details_2(5,array_add_new) = var_preorder_text
		array_details_2(6,array_add_new) = rsGetInfo("ProductID")
		array_details_2(7,array_add_new) = "" ' item notes
		array_details_2(8,array_add_new) = 0
		array_details_2(9,array_add_new)= rsGetInfo("picture")
		array_details_2(10,array_add_new) = array_length
		array_details_2(11,array_add_new) = array_detail
		array_details_2(12,array_add_new) = rsGetInfo("free") 

		GetOrderItems = array_details_2
		response.write "items found<br>"
		
	'================================================================================================
	' END store details into a dynamic multidimensional array

End Function
%>