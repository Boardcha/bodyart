<%
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET noscript = 1 WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,rsGetInvoiceNum.Fields.Item("ID").Value))
	objCmd.Execute()
	
%>