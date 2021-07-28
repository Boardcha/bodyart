<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<link href="/CSS/baf.min.css?v=120318" rel="stylesheet" type="text/css" />
<%
' =================== REQUEST SHIPPING LABEL 
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

'  Connect to the REST server.
bTls = 1
port = 443
bAutoReconnect = 1
success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
success = rest.ClearAllQueryParams()

success = rest.AddQueryParam("manifestId", "a07911c9-f3a7-4eed-8c5b-d68d83e90ab6")


' Set the Authorization property to "Bearer <token>"
	set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sbAuthHeaderVal.Append("Bearer ")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
	rest.Authorization = sbAuthHeaderVal.GetAsString()

ResponseGetTrack = rest.FullRequestNoBody("GET","/tracking/v4/package")

    set JsonTracking = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    JsonTracking.EmitCompact = 0
    JsonTracking.Load(ResponseGetTrack)
    'Response.Write "<pre>" & Server.HTMLEncode( JsonTracking.Emit()) & "</pre>"

if JsonTracking.IntOf("packages") <> "" then ' SUCCESS 

    if JsonTracking.StringOf("packages[0].package.expectedDelivery") <> "" then
        response.write FormatDateTime(JsonTracking.StringOf("packages[0].package.expectedDelivery"),1)
    end if

if JsonTracking.StringOf("packages[0].events") <> "" then ' ONLY LOOP IF THERE ARE EVENTS TO SHOW 

Set eventsArray = JsonTracking.ArrayOf("packages[0].events")
eventsSize = eventsArray.Size

j = eventsSize - 1 '---- REVERSES ORDER OF EVENTS FROM NEWEST TO OLDEST

    For e = 0 To eventsSize - 1
        
        Set eventsObj = eventsArray.ObjectAt(j)

        if eventsObj.StringOf("primaryEventDescription") <> "" then
        
            var_message = "<div class=""row py-2 border-top""><div class=""col-5"">" & FormatDateTime(replace(eventsObj.StringOf("date"), "-", "/"), 1) & "<br>" &  REPLACE(FormatDateTime(LEFT(eventsObj.StringOf("time"), 5), 3), ":00", "") & "<br>" &  eventsObj.StringOf("location") & "</div><div class=""col-7"">" &  eventsObj.StringOf("primaryEventDescription") & " " & eventsObj.StringOf("secondaryEventDescription")  & "</div></div>" & var_message 

        end if '--- only show if there is an event description

        '   "primaryEventId": 598
        '   "primaryEventDescription": "OUT FOR DELIVERY"

        '   "primaryEventId": 600
        if eventsObj.StringOf("primaryEventDescription") = "DELIVERED" then
            var_expected_delivery = ""
            var_delivered = "yes"
        end if

        j = j-1
    Next
    
end if ' only show if there are events
end if  ' Only if a PACKAGE is FOUND



%>