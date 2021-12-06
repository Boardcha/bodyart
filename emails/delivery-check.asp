<!--#include virtual="/Connections/sql_connection.asp" -->
<% Response.Buffer = False %>
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<!--#include virtual="/emails/function-send-email.asp"-->
<%
'=== CHECK ORDERS WILL BE DELIVERED TODAY ===
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE estimated_delivery_date = CONVERT(VARCHAR(10), GETDATE(), 23) AND delivering_today_email_sent = 0" 
Set rsDeliveringToday = objCmd.Execute()

While Not rsDeliveringToday.EOF 
	status = getDeliveryStatus(rsDeliveringToday("USPS_tracking"))
	If status = "OUT_FOR_DELIVERY" Then 
		mailer_type = "OUT_FOR_DELIVERY"
		var_email = "amanda@bodyartforms.com"
		'rsDeliveringToday("email")
		var_first = rsDeliveringToday("customer_first")
		var_invoiceid = rsDeliveringToday("ID")
			var_tracking = "Your tracking # is <strong>" & rsDeliveringToday("USPS_tracking") & "</strong>. If you have an account on our website, you can track your package by going to your order history and pressing the Track Order button. Or, you can track your package by going directly to <a href=""https://bodyartforms.com/dhl-tracker.asp?tracking=" & rsDeliveringToday("USPS_tracking") & """>this link</a>."		
		%>
		<!--#include virtual="/emails/email_variables.asp"-->
		<%
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET delivering_today_email_sent = 1 WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("packaged_delivered",3,1,15, rsDeliveringToday("ID")))
		objCmd.Execute()	
	End If
	rsDeliveringToday.MoveNext
Wend

'=== CHECK DELIVERED ORDERS ===
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE estimated_delivery_date = CONVERT(VARCHAR(10), GETDATE(), 23) AND delivered_email_sent = 0" 
Set rsDelivered = objCmd.Execute()

While Not rsDelivered.EOF 
	status = getDeliveryStatus(rsDelivered("USPS_tracking"))
	If status = "ORDER_DELIVERED" Then 
		mailer_type = "ORDER_DELIVERED"
		var_email = "amanda@bodyartforms.com"
		'rsDelivered("email")
		var_first = rsDelivered("customer_first")
		var_invoiceid = rsDelivered("ID")
		%>
		<!--#include virtual="/emails/email_variables.asp"-->
		<%
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET delivered_email_sent = 1 WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("packaged_delivered",3,1,15, rsDelivered("ID")))
		objCmd.Execute()	
	End If
	rsDelivered.MoveNext
Wend


'=== CHECK DELAYED ORDER ===
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE estimated_delivery_date = CONVERT(VARCHAR(10), DateAdd(""d"", - 1, GETDATE()), 23) AND packaged_delivered = 0 AND order_delayed_email_sent = 0" 
Set rsDelayed = objCmd.Execute()

While Not rsDelayed.EOF 
	mailer_type = "ORDER_DELAYED"
	var_email =  "amanda@bodyartforms.com"
	'rsDelayed("email")
	var_first = rsDelayed("customer_first")
	var_invoiceid = rsDelayed("ID")
	var_estimated_delivery_date = rsDelayed("estimated_delivery_date")
		var_tracking = "Your tracking # is <strong>" & rsDelayed("USPS_tracking") & "</strong>. If you have an account on our website, you can track your package by going to your order history and pressing the Track Order button. Or, you can track your package by going directly to <a href=""https://bodyartforms.com/dhl-tracker.asp?tracking=" & rsDelayed("USPS_tracking") & """>this link</a>."			
	%>
	<!--#include virtual="/emails/email_variables.asp"-->
	<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET order_delayed_email_sent = 1 WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("packaged_delivered",3,1,15, rsDelayed("ID")))
	objCmd.Execute()	
	rsDelayed.MoveNext
Wend

Set rsDeliveringToday = Nothing
Set rsDelivered = Nothing
Set rsDelayed = Nothing
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
	success = sbAuthHeaderVal.Append("Bearer")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
	rest.Authorization = sbAuthHeaderVal.GetAsString()

	ResponseGetTrack = rest.FullRequestNoBody("GET","/tracking/v4/package")
	set JsonTracking = Server.CreateObject("Chilkat_9_5_0.JsonObject")
	JsonTracking.EmitCompact = 0
	JsonTracking.Load(ResponseGetTrack)
	Response.Write "<pre>" & Server.HTMLEncode( JsonTracking.Emit()) & "</pre>"

	If JsonTracking.IntOf("packages") <> "" Then 

		If JsonTracking.StringOf("packages[0].events") <> "" Then ' ONLY LOOP If THERE ARE EVENTS TO SHOW 
			Set eventsArray = JsonTracking.ArrayOf("packages[0].events")
			eventsSize = eventsArray.Size
			j = eventsSize - 1

			For e = 0 To eventsSize - 1	
				Set eventsObj = eventsArray.ObjectAt(j)
				If eventsObj.IntOf("primaryEventId") = 598 Then
					var_status = "OUT_FOR_DELIVERY"
				End If

				If eventsObj.StringOf("primaryEventDescription") = "DELIVERED" Then
					var_status = "ORDER_DELIVERED"
					'======== WRITE DELIVERED STATUS TO DB ===================================
					set objCmd = Server.CreateObject("ADODB.command")
					objCmd.ActiveConnection = DataConn
					objCmd.CommandText = "UPDATE sent_items SET date_delivered = ?, packaged_delivered = ? WHERE USPS_tracking = ?"
					objCmd.Parameters.Append(objCmd.CreateParameter("date_delivered",135,1,30, eventsObj.StringOf("date") & " " & eventsObj.StringOf("time") ))
					objCmd.Parameters.Append(objCmd.CreateParameter("packaged_delivered",3,1,2, 1))
					objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking",200,1,200,request.querystring("tracking") ))
					objCmd.Execute()
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
%>