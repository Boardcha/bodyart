<%
If CustID_Cookie <> "" or var_our_custid <> "" Then ' if customer is logged in

' if paypal or cash or $0, do not save a billing profile ID (cuz there isn't one)
if request.form("cash") = "on" or request.form("cim_billing") = "cash" or request.form("paypal") = "on" or request.form("cim_billing") = "paypal" or var_grandtotal <= 0 then
	var_cim_billing_id = 0
end if

	' Update newest order to have all their CIM profiles
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET customer_ID = ?, cim_id = ?, shipping_profile_id = ? , payment_profile_id = ? WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("baf_custid",3,1,15, var_our_custid))
	objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",3,1,15, var_cim_custid))
	objCmd.Parameters.Append(objCmd.CreateParameter("shipping_profile_id",3,1,15,var_cim_shipping_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("payment_profile_id",3,1,15,var_cim_billing_id))	
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15, Session("invoiceid")))
	Set rsGetUser = objCmd.Execute()
	
end if

%>