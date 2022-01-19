
<%
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

%>