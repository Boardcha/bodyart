<%
'Response.Buffer = false
Server.ScriptTimeout = 2000
%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<%
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ID, USPS_tracking FROM sent_items WHERE international_tracking_num is null AND USPS_tracking is not null AND country <> 'US' AND country <> 'USA' AND country <> 'United States' AND shipping_type LIKE '%DHL%' AND date_sent >= CONVERT(VARCHAR(10), DateAdd(""d"", - 30, GETDATE()), 23)" 
Set rsTracking = objCmd.Execute()

While Not rsTracking.EOF
	trackingId = getInternationalTrackingId(rsTracking("USPS_tracking"))
	If trackingId <> "" Then 
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET international_tracking_num = '" & trackingId & "' WHERE ID = " & rsTracking("ID")
		objCmd.Execute()	
	End If
	rsTracking.MoveNext
Wend
rsTracking.Close

Response.Write "Successfuly completed." 
Set rsTracking = Nothing
DataConn.Close()
Set DataConn = Nothing
%>

<%
'=== FUNCTIONS ===
Function getInternationalTrackingId(trackingNumber)

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
		If JsonTracking.StringOf("packages[0].package.trackingId") <> "" Then 
			var_status = JsonTracking.StringOf("packages[0].package.trackingId")
		Else 
			var_status = "" 'No tracking information available yet"
		End If 
	Else ' NO PACKAGE FOUND
		var_status = "" 'No tracking information available yet"
	End If
	
	getInternationalTrackingId = var_status
	
End Function
%>