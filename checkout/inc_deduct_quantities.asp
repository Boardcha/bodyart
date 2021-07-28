<%
' Deduct quantities -- array is generated in inc_orderdetails_toarray.asp

	For i = 0 to (ubound(array_details_2, 2) - 1) ' loop through array
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - ?, DateLastPurchased = '" & date() & "' WHERE ProductDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("QtyPurchased",3,1,10,array_details_2(1,i)))		
		objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,12,array_details_2(0,i)))		
		objCmd.Execute()
		
	next ' loop through array
%>