<% 
' this line needs to remain under the google scripts include because it uses the google total variable

' Set an order to pending (bypass the review stage in the admin) only if one of these scenarios occur:
'- Order is below $150
'- Order does not have customer comments
'- Order does not have another order by the same customer submitted in the last hour
'- has a preorder on it
'- a problem city or area
'- Order has a gift certificate on it

'response.write "Form comments: " & request.form("Comments")
'response.write "<br>Session comments: " & session("customer_comments")
'response.write "<br>Grand total: " & var_grandtotal
'response.write "<br>City: " & session("city")
'response.write "<br>Preorder: " & var_grapreorder_shipping_noticendtotal

'======== Flag orders in database that are over $150 ===================================
if var_total_without_certsOrCredits > 150 then 
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET over_150 = 1 WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
	objCmd.Execute()
end if

' See if any orders have been placed in the last hour and if so, leave the order on review status
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ID FROM sent_items WHERE email = ? AND ship_code = 'paid' AND ID <> ? AND date_order_placed > DATEADD(HOUR, -1, GETDATE())"
objCmd.Parameters.Append(objCmd.CreateParameter("@Email",200,1,70,session("email")))
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
set rsLastHoursInvoices = objCmd.Execute()

if rsLastHoursInvoices.bof and rsLastHoursInvoices.eof then ' if no orders are found then push the order to pending status

if request.form("Comments") = "" and session("customer_comments") = "" and var_giftcert <> "yes" then 

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET shipped = 'Pending...' WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
		objCmd.Execute()

end if ' only run if all critera are met

if preorder_shipping_notice = "yes"  then ' If there's a preorder push it into review
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET shipped = 'PRE-ORDER REVIEW' WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
		objCmd.Execute()

end if  ' preorder_shipping_notice = "yes"

' Push orders through over $150 if they have more than 10 paid orders
if request.form("Comments") = "" and session("customer_comments") = "" then 

	' Count how many orders there are for email address with custID = 0
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID FROM sent_items WHERE email = ? AND customer_ID <> 0 AND ship_code = 'paid' AND ID <> ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("@Email",200,1,70,session("email")))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
	set rsLastHoursInvoices = Server.CreateObject("ADODB.Recordset")
	rsLastHoursInvoices.CursorLocation = 3 'adUseClient
	rsLastHoursInvoices.Open objCmd
	var_total_orders = rsLastHoursInvoices.RecordCount	
	
	if var_total_orders >= 10 then
	
		if preorder_shipping_notice <> "yes" then ' If there's no preorder in the order, set it to ship out
	
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE sent_items SET shipped = 'Pending...' WHERE ID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
			objCmd.Execute()

		else ' If order has a pre-order in it, set it to preorder review
		
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE sent_items SET shipped = 'PRE-ORDER REVIEW' WHERE ID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
			objCmd.Execute()
		
		end if  ' preorder_shipping_notice <> "yes"

			
	end if  ' var_total_orders >= 10
	
end if  ' var_grandtotal > 150



end if  '  rsLastHoursInvoices.eof
%>