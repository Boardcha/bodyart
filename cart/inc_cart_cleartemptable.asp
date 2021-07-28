<%

if Request.Cookies("ID") = "" then 
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "DELETE FROM tbl_carts_temp WHERE cart_sessionID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_sessionID",3,1,10,Session.SessionID))
		objCmd.Execute()
end if

%>