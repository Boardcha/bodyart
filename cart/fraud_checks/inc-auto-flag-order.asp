<%
remote_ip = Request.ServerVariables("REMOTE_HOST")
domainname = Request.ServerVariables("server_name")
' exclude DEV5 domain, georgetown and localhost ip addresses
if Instr(domainname, "dev") = 0 And remote_ip <> "::1" and remote_ip <> "75.109.218.250" and remote_ip <> "75.109.218.58" and remote_ip <> "70.114.165.125" and remote_ip <> "127.0.0.1" _
	And remote_ip <> "49.255.77.150" And remote_ip <> "49.255.76.138" And remote_ip <> "115.186.198.66" And remote_ip <> "124.19.8.202" And remote_ip <> "78.27.129.3" And remote_ip <> "185.190.151.100" And remote_ip <> "18.215.8.102" And remote_ip <> "52.52.23.83" And remote_ip <> "3.104.192.118" then 'IPs in this line are for Afterpay

	' Search ALL orders for anything matching current IP address and return 6 results as the threshold to auto flag
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP 15 ID, IPaddress FROM sent_items WHERE date_order_placed > DATEADD(HOUR, -12, GETDATE()) AND IPaddress = ? ORDER BY ID DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("ip_addy",200,1,40,remote_ip))


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