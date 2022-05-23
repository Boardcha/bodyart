<%
if Request.ServerVariables("SERVER_NAME") <> "dev5.bodyartforms.com" and Request.ServerVariables("SERVER_NAME") <> "localhost" and Request.ServerVariables("REMOTE_HOST") <> "::1" and Request.ServerVariables("REMOTE_HOST") <> "75.109.218.250" and Request.ServerVariables("REMOTE_HOST") <> "75.109.218.58" and Request.ServerVariables("REMOTE_HOST") <> "70.114.165.125" and Request.ServerVariables("REMOTE_HOST") <> "127.0.0.1"  then ' exclude georgetown and localhost ip address

' Search ALL orders for anything matching current IP address and return 6 results as the threshold to auto flag
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP 15 ID, IPaddress FROM sent_items WHERE date_order_placed > DATEADD(HOUR, -12, GETDATE()) AND IPaddress = ? ORDER BY ID DESC"
objCmd.Parameters.Append(objCmd.CreateParameter("ip_addy",200,1,40,Request.ServerVariables("REMOTE_HOST")))


set rs_getFlaggedOrders = Server.CreateObject("ADODB.Recordset")
rs_getFlaggedOrders.CursorLocation = 3 'adUseClient
rs_getFlaggedOrders.Open objCmd
var_total_orders = rs_getFlaggedOrders.RecordCount

if not rs_getFlaggedOrders.eof and var_total_orders > 6  then

	session("flag") = "yes"
	mailer_type = "auto-flag"
	invoice_id = rs_getFlaggedOrders.Fields.Item("ID").Value
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET shipped = 'FLAGGED', our_notes = 'Website trip auto-flagged this order. More than 6 orders were attempted to be placed within a 12 hour period.' WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, rs_getFlaggedOrders.Fields.Item("ID").Value))
	objCmd.Execute()
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->

<%	
end if

end if ' exclude georgetown ip address

' Unflag the order for testing purposes so we can run orders back to back
'session("flag") = ""
%>
