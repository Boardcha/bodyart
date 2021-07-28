<%
' Pull all flagged orders and loop through them to set a flagged status for page
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT email, IPaddress FROM sent_items WHERE (shipped = 'FLAGGED' or shipped = 'Chargeback') AND IPaddress = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("ip_addy",200,1,40,Request.ServerVariables("REMOTE_HOST")))
Set rs_getFlaggedOrders = objCmd.Execute()

if not rs_getFlaggedOrders.eof then
	Flagged = "yes"
end if
%>