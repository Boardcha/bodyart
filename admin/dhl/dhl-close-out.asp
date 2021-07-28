<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
{
<%
' =================== SUBMIT CLOSE OUT FORM   
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

'  Connect to the REST server.
bTls = 1
port = 443
bAutoReconnect = 1
success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
success = rest.AddHeader("Content-Type","application/json")

	set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sbAuthHeaderVal.Append("Bearer ")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
    rest.Authorization = sbAuthHeaderVal.GetAsString()
    
ResponseRequestCloseout = rest.FullRequestString("POST","/shipping/v4/manifest", "{""pickup"": """ & dhl_production_pickup_num & """,""manifests"":[]}")

    set JsonCloseout = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    JsonCloseout.EmitCompact = 0
    JsonCloseout.Load(ResponseRequestCloseout)
    'Response.Write "<pre>" & Server.HTMLEncode( JsonCloseout.Emit()) & "</pre>"

if JsonCloseout.StringOf("requestId") <> "" then ' SUCCESS 
    var_request_id = JsonCloseout.StringOf("requestId")
  
    ' Store closeouts in the database
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "INSERT INTO tbl_closeout_forms (provider, requestId) VALUES ('DHL', ?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("requestId",200,1,1000, var_request_id))
    objCmd.Execute()

%>
    "status":"success",
    "message": "success"
<%
else

%>
    "status":"error",
    "message": "<%= JsonCloseout.StringOf("invalidParams[0].reason") %>"
<%
end if


%>
}